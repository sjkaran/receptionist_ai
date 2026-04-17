import anthropic

# Read API key from file
with open("APIKEY.txt", "r") as f:
    api_key = f.read().strip()

client = anthropic.Anthropic(
    api_key=api_key,
    base_url="https://api.minimax.io/anthropic"
)

try:
    message = client.messages.create(
        model="MiniMax-M2.7",
        max_tokens=100,
        system="You are a helpful assistant.",
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": "Hello"
                    }
                ]
            }
        ]
    )

    print("Success!")
    for block in message.content:
        print(block)
except Exception as e:
    print(f"Error: {e}")