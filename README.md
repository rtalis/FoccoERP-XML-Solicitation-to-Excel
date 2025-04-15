# XML to XLSX Converter

Este script importa um arquivo XML do FoccoERP - Solicitação de Compra e cria um arquivo XLSX com os itens a serem precificados. O arquivo Excel gerado contém as descrições e quantidades dos itens.

## Funcionalidades

- Lê dados XML do FoccoERP Solicitação de Compra.
- Extrai descrições e quantidades dos itens.
- Salva os dados em um arquivo XLSX com um nome de arquivo com timestamp.

## Requisitos

- Python 3.x
- Biblioteca `openpyxl`
- Biblioteca `tkinter` (para diálogo de arquivos)

## Configuração

Instale as bibliotecas necessárias:

```bash
pip install -r requirements.txt

python main.py <xml_file>

```
## Uso
Execute o script com o arquivo XML como argumento:
```bash
python main.py <xml_file>
```
Alternativamente, você pode arrastar e soltar o arquivo XML no executável (uma vez criado).

## Criando um Executável
Para criar um executável standalone, use o PyInstaller.
Use o seguinte comando para criar o executável:
```bash
pyinstaller --onefile --add-data="model1.xlsx:." --add-data="model2.xlsx:." main.py
```
Este comando empacota o script e o arquivo model.xlsx em um único arquivo executável. 

## Exemplo
Para processar um arquivo XML chamado data.xml:
```bash
python main.py RPDC0251.xml
```