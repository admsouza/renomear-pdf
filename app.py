from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import os
import pandas as pd
from werkzeug.utils import secure_filename
import zipfile
import io

app = Flask(__name__)
app.secret_key = "secret_key"

def renomear_arquivos(arquivos_pdfs, planilha_caminho):
    df = pd.read_excel(planilha_caminho)

    # Verificar as colunas
    required_columns = ['De', 'Para']
    if not all(column in df.columns for column in required_columns):
        flash("Erro: A planilha deve ter as colunas 'De' e 'Para'.", 'error')
        return None

    # Criar um buffer de memória para o ZIP
    zip_buffer = io.BytesIO()
    
    # Criar o arquivo ZIP na memória
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        # Renomear arquivos e adicioná-los ao ZIP
        for index, row in df.iterrows():
            nome_atual = f"{row['De']}.pdf"
            novo_nome = f"{row['Para']}.pdf"

            # Procurar o arquivo enviado correspondente
            for arquivo in arquivos_pdfs:
                if secure_filename(arquivo.filename) == nome_atual:
                    # Adicionar o arquivo renomeado ao ZIP
                    zip_file.writestr(novo_nome, arquivo.read())
                    break
            else:
                flash(f"Arquivo {nome_atual} não encontrado nos arquivos enviados.", 'error')

    zip_buffer.seek(0)  # Posicionar o buffer no início
    return zip_buffer

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        arquivos_pdfs = request.files.getlist("arquivos_pdfs")
        planilha = request.files["planilha"]

        # Verificar se os arquivos foram enviados
        if not arquivos_pdfs or not planilha:
            flash("Por favor, forneça a planilha e os arquivos PDFs.", 'error')
            return redirect(url_for('index'))

        # Salvar a planilha temporariamente no buffer
        planilha_caminho = os.path.join(os.getcwd(), secure_filename(planilha.filename))
        planilha.save(planilha_caminho)

        # Gerar arquivo ZIP com PDFs renomeados
        zip_buffer = renomear_arquivos(arquivos_pdfs, planilha_caminho)

        if zip_buffer:
            # Enviar o arquivo ZIP para o usuário
            return send_file(
                zip_buffer,
                as_attachment=True,
                download_name="arquivos_renomeados.zip",
                mimetype="application/zip"
            )

    return render_template("index.html")

if __name__ == "__main__":
    # Definir a porta para o ambiente de produção ou usar a porta 5000 por padrão
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
