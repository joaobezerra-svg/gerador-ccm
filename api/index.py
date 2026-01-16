from flask import Flask

app = Flask(__name__)

@app.route('/', defaults={'path': ''})
@app.route('/<path:path>')
def catch_all(path):
    return "<h1>Sistema Online! (Modo de Teste)</h1><p>Se você está vendo isso, o Vercel está funcionando e o erro estava no código Python.</p>"

# Se precisar voltar o código original, reverta este commit.
