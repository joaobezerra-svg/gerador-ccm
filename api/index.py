from flask import Flask

app = Flask(__name__) # O Vercel procura por esta linha exata

@app.route('/')
def home():
    return "Servidor Online"

# Não precisa de app.run() no Vercel
