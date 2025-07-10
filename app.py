import os
import pandas as pd
from datetime import datetime
from flask import Flask, request, jsonify, render_template, send_file

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def uploadFile():
    if "arquivo" not in request.files:
        return jsonify({"Error": "Nenhum arquivo enviado"}), 400
    
    arquivo = request.files["arquivo"]

    if arquivo.filename == "":
        return jsonify({"Error": "Nenhum arquivo selecionado."}), 400
    
    if not arquivo.filename.endswith(".xlsx"):
        return jsonify({"Error": "Formato de arquivo inválido."}), 400 
    
    # caminho onde o arquivo será salvo
    caminho_arquivo = os.path.join(UPLOAD_FOLDER, arquivo.filename)
    arquivo.save(caminho_arquivo)

    try:
        df = pd.read_excel(caminho_arquivo)
        df["Data"] = pd.to_datetime(df["Data"])

        # Captura dos filtros (se houver)
        data_inicio_str = request.form.get("data_inicio")
        data_fim_str = request.form.get("data_fim")

        # Conversão para datetime
        data_inicio = datetime.strptime(data_inicio_str, "%Y-%m-%d") if data_inicio_str else None
        data_fim = datetime.strptime(data_fim_str, "%Y-%m-%d") if data_fim_str else None

        # Aplicação dos filtros
        if data_inicio and data_fim:
            df = df[(df["Data"] >= data_inicio) & (df["Data"] <= data_fim)]
        elif data_inicio:
            df = df[df["Data"] >= data_inicio]
        elif data_fim:
            df = df[df["Data"] <= data_fim]

        # Ordenação por Valor (ou Data, se preferir)
        df = df.sort_values(by="Valor", ascending=True)

        # Gera novo Excel
        novo_arquivo = os.path.join(UPLOAD_FOLDER, "relatorio_filtrado.xlsx")
        df.to_excel(novo_arquivo, index=False)

        return render_template("result.html", filename=arquivo.filename)
    
    except Exception as e:
        return jsonify({"Error": f"Erro ao processar o arquivo: {str(e)}"}), 500


@app.route("/download")
def download():
    caminho_arquivo = os.path.join(UPLOAD_FOLDER, "relatorio_filtrado.xlsx")
    return send_file(caminho_arquivo, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)