using System;
using System.IO;
using System.Net.WebSockets;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

class OBSControl
{
    static string host = "localhost";
    static int port = 4455;
    static string password = "";
    static string sceneName = "Scene";
    static string sourceName = "Video Capture Device 2";

    static void Main(string[] args)
    {
        // Parse arguments
        for (int i = 0; i < args.Length; i++)
        {
            if (args[i] == "--password" && i + 1 < args.Length)
                password = args[++i];
            else if (args[i] == "--scene" && i + 1 < args.Length)
                sceneName = args[++i];
            else if (args[i] == "--source" && i + 1 < args.Length)
                sourceName = args[++i];
            else if (args[i] == "--port" && i + 1 < args.Length)
                port = int.Parse(args[++i]);
            else if (args[i] == "--help" || args[i] == "-h")
            {
                PrintUsage();
                return;
            }
        }

        try
        {
            RestartSource().Wait();
            Console.WriteLine("Source restarted successfully!");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
            Environment.Exit(1);
        }
    }

    static void PrintUsage()
    {
        Console.WriteLine("OBSControl - Restart OBS source via WebSocket");
        Console.WriteLine();
        Console.WriteLine("Usage: OBSControl.exe [options]");
        Console.WriteLine();
        Console.WriteLine("Options:");
        Console.WriteLine("  --password <pass>   OBS WebSocket password");
        Console.WriteLine("  --scene <name>      Scene name (default: Scene)");
        Console.WriteLine("  --source <name>     Source name (default: Video Capture Device 2)");
        Console.WriteLine("  --port <port>       WebSocket port (default: 4455)");
    }

    static async Task RestartSource()
    {
        using (var ws = new ClientWebSocket())
        {
            var uri = new Uri("ws://" + host + ":" + port);
            Console.WriteLine("Connecting to OBS WebSocket...");

            var cts = new CancellationTokenSource(5000);
            await ws.ConnectAsync(uri, cts.Token);
            Console.WriteLine("Connected!");

            // Receive Hello message
            string hello = await ReceiveMessage(ws);
            Console.WriteLine("Received Hello");

            // Parse authentication challenge if present
            string authChallenge = ExtractJsonValue(hello, "challenge");
            string authSalt = ExtractJsonValue(hello, "salt");

            // Send Identify message
            string authString = "";
            if (!string.IsNullOrEmpty(password) && !string.IsNullOrEmpty(authChallenge))
            {
                authString = GenerateAuthString(password, authChallenge, authSalt);
            }

            string identify;
            if (!string.IsNullOrEmpty(authString))
            {
                identify = "{\"op\":1,\"d\":{\"rpcVersion\":1,\"authentication\":\"" + authString + "\"}}";
            }
            else
            {
                identify = "{\"op\":1,\"d\":{\"rpcVersion\":1}}";
            }

            await SendMessage(ws, identify);
            string identifyResponse = await ReceiveMessage(ws);

            if (identifyResponse.Contains("\"op\":2"))
            {
                Console.WriteLine("Authenticated!");
            }
            else
            {
                throw new Exception("Authentication failed. Check password.");
            }

            // Get scene item ID
            Console.WriteLine("Getting scene item ID for '" + sourceName + "'...");
            string getItemId = "{\"op\":6,\"d\":{\"requestType\":\"GetSceneItemId\",\"requestId\":\"1\",\"requestData\":{\"sceneName\":\"" + sceneName + "\",\"sourceName\":\"" + sourceName + "\"}}}";
            await SendMessage(ws, getItemId);
            string itemIdResponse = await ReceiveMessage(ws);

            string sceneItemId = ExtractJsonValue(itemIdResponse, "sceneItemId");
            if (string.IsNullOrEmpty(sceneItemId))
            {
                throw new Exception("Could not find source '" + sourceName + "' in scene '" + sceneName + "'");
            }
            Console.WriteLine("Scene item ID: " + sceneItemId);

            // Disable source
            Console.WriteLine("Disabling source...");
            string disableSource = "{\"op\":6,\"d\":{\"requestType\":\"SetSceneItemEnabled\",\"requestId\":\"2\",\"requestData\":{\"sceneName\":\"" + sceneName + "\",\"sceneItemId\":" + sceneItemId + ",\"sceneItemEnabled\":false}}}";
            await SendMessage(ws, disableSource);
            await ReceiveMessage(ws);

            // Wait for driver to reinitialize with new input
            Console.WriteLine("Waiting 3 seconds for input switch...");
            await Task.Delay(3000);

            // Enable source
            Console.WriteLine("Enabling source...");
            string enableSource = "{\"op\":6,\"d\":{\"requestType\":\"SetSceneItemEnabled\",\"requestId\":\"3\",\"requestData\":{\"sceneName\":\"" + sceneName + "\",\"sceneItemId\":" + sceneItemId + ",\"sceneItemEnabled\":true}}}";
            await SendMessage(ws, enableSource);
            await ReceiveMessage(ws);

            await ws.CloseAsync(WebSocketCloseStatus.NormalClosure, "", CancellationToken.None);
        }
    }

    static async Task SendMessage(ClientWebSocket ws, string message)
    {
        var bytes = Encoding.UTF8.GetBytes(message);
        await ws.SendAsync(new ArraySegment<byte>(bytes), WebSocketMessageType.Text, true, CancellationToken.None);
    }

    static async Task<string> ReceiveMessage(ClientWebSocket ws)
    {
        var buffer = new byte[8192];
        var result = await ws.ReceiveAsync(new ArraySegment<byte>(buffer), CancellationToken.None);
        return Encoding.UTF8.GetString(buffer, 0, result.Count);
    }

    static string GenerateAuthString(string password, string challenge, string salt)
    {
        // OBS WebSocket 5.x auth: base64(sha256(base64(sha256(password + salt)) + challenge))
        using (var sha256 = SHA256.Create())
        {
            // First hash: sha256(password + salt)
            byte[] secretBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(password + salt));
            string secret = Convert.ToBase64String(secretBytes);

            // Second hash: sha256(secret + challenge)
            byte[] authBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(secret + challenge));
            return Convert.ToBase64String(authBytes);
        }
    }

    static string ExtractJsonValue(string json, string key)
    {
        // Simple JSON value extractor
        string searchKey = "\"" + key + "\":";
        int keyIndex = json.IndexOf(searchKey);
        if (keyIndex < 0) return "";

        int valueStart = keyIndex + searchKey.Length;

        // Skip whitespace
        while (valueStart < json.Length && char.IsWhiteSpace(json[valueStart]))
            valueStart++;

        if (valueStart >= json.Length) return "";

        // Check if it's a string value
        if (json[valueStart] == '"')
        {
            int valueEnd = json.IndexOf('"', valueStart + 1);
            if (valueEnd > valueStart)
                return json.Substring(valueStart + 1, valueEnd - valueStart - 1);
        }
        // Check if it's a number
        else if (char.IsDigit(json[valueStart]) || json[valueStart] == '-')
        {
            int valueEnd = valueStart;
            while (valueEnd < json.Length && (char.IsDigit(json[valueEnd]) || json[valueEnd] == '.' || json[valueEnd] == '-'))
                valueEnd++;
            return json.Substring(valueStart, valueEnd - valueStart);
        }

        return "";
    }
}
