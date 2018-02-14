from Solubility import app, solubility
       
@app.route('/')
@app.route('/index')
def index():
    solubility.saveFile()
    return "Hello, World!"