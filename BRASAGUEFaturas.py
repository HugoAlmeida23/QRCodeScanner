import PySimpleGUI as sg
import cv2
from qreader import QReader
import pandas as pd

qreader = QReader()
sg.theme("DarkGrey5")
    
def parse_qr_text(qr_text):
    print("QR Text:", qr_text)  # Adicionando esta linha para depuração
    parsed_data = {}
    if qr_text is not None:
        fields = qr_text.split("*")
        for field in fields:
            print("Field:", field)  # Adicionando esta linha para depuração
            if ':' in field:  # Verificando se há ':' na string antes de dividir
                key, value = field.split(":", 1)  # Dividindo a string apenas uma vez
                parsed_data[key] = value
        return parsed_data
    else:
        return {}  

def update_image_and_data(window, filepath, filepathexcel):
    image = cv2.imread(filepath)
    image = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
    image = cv2.resize(image, (400, 200))
    window["-IMAGE-"].update(data=cv2.imencode('.png', image)[1].tobytes())
    df = pd.read_excel(filepathexcel)

    decoded_text = qreader.detect_and_decode(image=cv2.cvtColor(cv2.imread(filepath, 0), cv2.COLOR_BAYER_BG2BGR))

    print(decoded_text)
    print(len(decoded_text))
    print(decoded_text[0])
    if decoded_text[0] is not None:
    #if len(decoded_text[0]) > 15:
            if decoded_text and isinstance(decoded_text, tuple) and len(decoded_text) > 0 and len(decoded_text) < 15:
                qr_text = decoded_text[0]

                parsed_data = parse_qr_text(qr_text)
                parsed_data = {key: str(value) for key, value in parsed_data.items()}

                ## preencher o ecrã com os dados da fatura lida
                window["-N_Fatura-"].update(parsed_data.get("G", "N/A"))
                window["-NIF_Fornecedor-"].update(parsed_data.get("A", "N/A"))
                window["-NIF_Cliente-"].update(parsed_data.get("B", "N/A"))
                window["-Data-"].update(parsed_data.get("F", "N/A"))
                window["-Total-"].update(parsed_data.get("O", "N/A"))
                window["-IVA 1-"].update(parsed_data.get("I6", "N/A"))
                window["-IVA 2-"].update(parsed_data.get("I8", "N/A"))

                # Inicializando uma lista para armazenar as diferenças
                flag = 0

                # Obtendo o número da fatura extraído do QR code
                qr_faturanumber = parsed_data.get("G")

                # Percorrendo todas as linhas do DataFrame
                for index, row in df.iterrows():
                    # Obtendo o número da fatura da linha atual
                    faturanumber = row['Número']
                    print(faturanumber)
                    print(qr_faturanumber)
                    # Verificando se há uma diferença entre o número da fatura extraído do QR code e o número da fatura na linha atual
                    if str(faturanumber) == str(qr_faturanumber):
                        flag = 1

                if flag == 0:
                    #sg.popup("Os campos do QR code não correspondem aos dados no arquivo Excel. Diferenças encontradas nos seguintes campos:\n{}".format("\n".join(differences)))
                    print("nao encontrou nada")
                else:
                    sg.popup("Os campos do QR code correspondem aos dados no arquivo Excel.")
                    print("encontrou algo")
                    index_of_match = df[df['Número'].astype(str) == parsed_data.get("G")].index

                    if not index_of_match.empty:
                        df.loc[index_of_match, "Match"] = "Fatura encontrada"
                        df.to_excel(filepathexcel, index=False)
                        sg.popup("Fatura encontrada e arquivo atualizado com sucesso.") 
                        show_updated_data_window(df)
                        #show_updated_data_window(df)
                    else:
                        sg.popup("Os campos do QR code correspondem aos dados no arquivo Excel, mas não foi possível encontrar a linha correspondente para atualizar.")
    else:
        sg.popup("QR code não encontrado!")
    
   
 
def show_updated_data_window(df):
    layout = [
        [sg.Text("Dados atualizados:", font=("Helvetica", 20))],
        [sg.Column([
            [sg.Table(values=df.values.tolist(), 
                      headings=df.columns.tolist(), 
                      display_row_numbers=False, 
                      auto_size_columns=True, 
                      num_rows=min(25, len(df)),
                      vertical_scroll_only=False)]  # Enable horizontal scrolling
        ], expand_y=True)]  # Expand vertically
    ]
    window = sg.Window("Dados Atualizados", layout, modal=True, resizable=True)
    event, _ = window.read()
    window.close()

def excelFile():
    firstlayout = [
        [sg.Button("Escolher Ficheiro Excel", size=(15,2), pad=(10,10))],
        [sg.Button("Sair",size=(10,1), pad=(10,10))],
    ]
    
    firstwindow = sg.Window("Seleção Ficheiro Excel!", firstlayout, margins=(20,20))
    
    while True:
        event, values = firstwindow.read()
        if event == sg.WIN_CLOSED or event == "Sair":
            break
        elif event == "Escolher Ficheiro Excel":
            filepathexcel = sg.popup_get_file('Selecione o ficheiro Excel!', no_window=True, file_types=(("ALL Files",".*"),))
            flag = main(filepathexcel)
            if flag == 1:
                break

def main(filepathexcel):
    layout = [
        [sg.Button("Upload Fatura", size=(15, 2), pad=(10, 10))],
        [sg.Image(key="-IMAGE-", size=(400, 200), pad=(10, 10))],
        [sg.Text("Nº da Fatura :", size=(15, 1), pad=(10, (5, 5))), sg.Text(size=(30, 1), key="-N_Fatura-")],
        [sg.Text("NIF do Fornecedor :", size=(15, 1), pad=(10, (5, 5))), sg.Text(size=(30, 1), key="-NIF_Fornecedor-")],
        [sg.Text("NIF do Cliente:", size=(15, 1), pad=(10, (5, 5))), sg.Text(size=(30, 1), key="-NIF_Cliente-")],
        [sg.Text("Data:", size=(15, 1), pad=(10, (5, 5))), sg.Text(size=(30, 1), key="-Data-")],
        [sg.Text("Total:", size=(15, 1), pad=(10, (5, 5))), sg.Text(size=(30, 1), key="-Total-")],
        [sg.Text("IVA 1:", size=(15, 1), pad=(10, (5, 5))), sg.Text(size=(30, 1), key="-IVA 1-")],
        [sg.Text("IVA 2:", size=(15, 1), pad=(10, (5, 5))), sg.Text(size=(30, 1), key="-IVA 2-")],
    ]
    
    window = sg.Window("BRASAGUE Faturas", layout, margins=(20, 20),resizable=True)
    sg.popup("Certifique-se que o ficheiro Excel esteja fechado antes de qualquer operação!!")
    flag = 0

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED:
            flag = 1
            break
        elif event == "Upload Fatura":
            filepaths = sg.popup_get_file('Selecione as imagens', no_window=True, file_types=(("Images Files", "*.png;*.jpg;*.jpeg"),), multiple_files=True)
            if filepaths:
                for filepath in filepaths:
                    update_image_and_data(window, filepath, filepathexcel)
    
    window.close()
    return flag

if __name__ == "__main__":
    excelFile()
