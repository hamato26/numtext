import pandas as pd
from tkinter import Tk, filedialog, Label, Button, Entry, Text, Scrollbar, Frame
from num2words import num2words

def convertir_decimal(decimal):
    if decimal == 0:
        return "  00 /100 soles"
    else:
        # Multiplicar por 100 y convertir a entero para obtener los primeros dos dígitos decimales
        parte_decimal = int(decimal * 100)
        return f" {num2words(parte_decimal, lang='es')} /100 soles"

def convertir_a_texto(precio):
    parte_entera = int(precio)
    parte_decimal = round(precio - parte_entera, 2)

    texto_entero = num2words(parte_entera, lang='es').upper()

    # Verificar si hay parte decimal
    if parte_decimal == 0:
        texto_decimal = " 00 /100 SOLES"
    else:
        # Multiplicar por 100 para obtener los primeros dos dígitos decimales
        parte_decimal = int(parte_decimal * 100)
        texto_decimal = f" CON {parte_decimal:02d} /100 SOLES"

    # Agregar "CON" solo para números decimales y cuando hay parte decimal diferente de 0
    if parte_decimal > 0 and parte_decimal % 10 == 0:
        texto_completo = f'{texto_entero} CON {parte_decimal // 10}0 /100 SOLES'
    else:
        texto_completo = f'{texto_entero}{"" if parte_entera == 0 else " CON"}{texto_decimal}'

    # Eliminar "CON" repetido
    texto_completo = texto_completo.replace("CON CON", "CON")

    return texto_completo




def procesar_archivo(ruta_archivo_entry, ubicacion_text):
    ruta_archivo = ruta_archivo_entry.get()
    
    if ruta_archivo:
        df = pd.read_excel(ruta_archivo)
        df.set_index(df.columns[0], inplace=True)

        df['Texto'] = df.index.map(convertir_a_texto)

        nuevo_nombre = ruta_archivo.replace('.xlsx', '_modificado.xlsx')
        df.to_excel(nuevo_nombre, index=True)
        ubicacion_text.config(state="normal")
        ubicacion_text.delete(1.0, "end")
        ubicacion_text.insert("insert", f"Archivo procesado y guardado como {nuevo_nombre}")
        ubicacion_text.config(state="disabled")

def seleccionar_archivo(ruta_archivo_entry):
    ruta_archivo = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx;*.xls")])
    ruta_archivo_entry.delete(0, "end")
    ruta_archivo_entry.insert("insert", ruta_archivo)

def main():
    root = Tk()
    root.title("Conversor de Precios a Texto")
    root.geometry("540x340")
    root.resizable(False, False)

    frame = Frame(root, bg="#E0E0E0")
    frame.pack(pady=20)

    label = Label(frame, text="Conversor de Precios a Texto", font=("Arial", 16), bg="#E0E0E0")
    label.grid(row=0, column=0, columnspan=3, pady=(10, 20))

    etiqueta_archivo = Label(frame, text="Seleccionar Archivo:", font=("Arial", 12), bg="#E0E0E0")
    etiqueta_archivo.grid(row=1, column=0, padx=20, pady=(0, 10), sticky="e")

    ruta_archivo_entry = Entry(frame, width=30, font=("Arial", 10))
    ruta_archivo_entry.grid(row=1, column=1, padx=5, pady=(0, 10), sticky="w")

    boton_seleccion = Button(frame, text="Seleccionar", command=lambda: seleccionar_archivo(ruta_archivo_entry), font=("Arial", 10), bg="#4CAF50", fg="white", padx=10, pady=5)
    boton_seleccion.grid(row=1, column=2, padx=(5, 20), pady=(0, 10), sticky="e")

    boton_conversion = Button(frame, text="Convertir", command=lambda: procesar_archivo(ruta_archivo_entry, ubicacion_text), font=("Arial", 12), bg="#4CAF50", fg="white", padx=10, pady=5)
    boton_conversion.grid(row=2, column=1, pady=(0, 20))

    ubicacion_text = Text(frame, height=3, width=40, font=("Arial", 10), wrap="word", state="disabled")
    ubicacion_text.grid(row=3, column=0, columnspan=3, padx=20, pady=10, sticky="w")

    ubicacion_scrollbar = Scrollbar(frame, command=ubicacion_text.yview)
    ubicacion_scrollbar.grid(row=3, column=3, pady=10, sticky="ns")

    ubicacion_text.config(yscrollcommand=ubicacion_scrollbar.set)

    root.mainloop()

if __name__ == "__main__":
    main()
