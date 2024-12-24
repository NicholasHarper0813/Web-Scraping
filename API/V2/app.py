
from flask import jsonify, Flask
import main


# start flask
app: Flask = Flask(__name__)

# route to get all datas
@app.route('/api/<id>', methods=["GET"])
def get_datas(id:str) -> jsonify:
    if id:
        datas: dict = main.Goaloo(id).start()
        return jsonify(datas)
    else:
        return jsonify({"erro": 'No ID'})

if __name__ == "__main__":
    app.run(debug=True, port=5500)