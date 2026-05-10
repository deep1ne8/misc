# Define the port to listen on
$port = 8080

# Create a listener
$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add("http://*:8080/")
$listener.Start()
Write-Output "HTTP server started. Listening on port $port..."

# Handle incoming requests
while ($listener.IsListening) {
    $context = $listener.GetContext()
    # $request = $context.Request
    $response = $context.Response

    # Define the response content
    $responseString = "<html><body>Hello, World!</body></html>"
    $buffer = [System.Text.Encoding]::UTF8.GetBytes($responseString)
    $response.ContentLength64 = $buffer.Length
    $response.OutputStream.Write($buffer, 0, $buffer.Length)
    $response.OutputStream.Close()
}

# Stop the listener when done
$listener.Stop()
$listener.Close()