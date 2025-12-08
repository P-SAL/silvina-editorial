import ollama

# Simple test to verify Ollama connection
response = ollama.chat(
    model='llama3.2:1b',
    messages=[
        {
            'role': 'user',
            'content': 'Responde en español: ¿Puedes revisar textos académicos?'
        }
    ]
)

print(response['message']['content'])