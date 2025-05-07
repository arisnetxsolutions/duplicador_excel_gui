import os
import threading
from tkinter import (
    Tk,
    filedialog,
    messagebox,
    Label,
    Entry,
    Button,
    StringVar,
    OptionMenu,
    IntVar,
    Checkbutton,
)
import pandas as pd
import subprocess


def letra_a_indice(letra):
    return ord(letra.upper()) - ord("A")


def procesar_excel_con_pandas(
    archivo_excel,
    nombre_hoja,
    letra_columna,
    tiene_header=False,
    letra_columna_extra=None,
):
    df = pd.read_excel(
        archivo_excel, sheet_name=nombre_hoja, header=0 if tiene_header else None
    )

    col_idx = letra_a_indice(letra_columna)
    col_extra_idx = letra_a_indice(letra_columna_extra) if letra_columna_extra else None

    if col_idx >= len(df.columns):
        raise ValueError("Índice de columna principal fuera de rango.")
    if col_extra_idx is not None and col_extra_idx >= len(df.columns):
        raise ValueError("Índice de columna extra fuera de rango.")

    col = df.columns[col_idx]
    col_extra = df.columns[col_extra_idx] if col_extra_idx is not None else None

    resultado = []
    i = 0
    while i < len(df):
        if pd.notna(df.at[i, col]) and str(df.at[i, col]).strip() != "":
            inicio_grupo = i
            while (
                i < len(df)
                and pd.notna(df.at[i, col])
                and str(df.at[i, col]).strip() != ""
            ):
                i += 1
            fin_grupo = i

            inicio_espacios = i
            while i < len(df) and (
                pd.isna(df.at[i, col]) or str(df.at[i, col]).strip() == ""
            ):
                i += 1
            fin_espacios = i

            espacios = fin_espacios - inicio_espacios
            if espacios > 0:
                bloque = df.iloc[inicio_grupo:fin_grupo]
                for _ in range(espacios):
                    resultado.append(bloque.copy())
        else:
            i += 1

    for bloque in reversed(resultado):
        idx = bloque.index[0]
        df = pd.concat([df.iloc[:idx], bloque, df.iloc[idx:]], ignore_index=True)

    if col_extra:
        df = df[df[col_extra].notna() & (df[col_extra].astype(str).str.strip() != "")]

    base, ext = os.path.splitext(archivo_excel)
    nuevo_nombre = f"{base}_duplicado.xlsx"
    df.to_excel(nuevo_nombre, index=False, sheet_name=nombre_hoja)
    return nuevo_nombre


# -------------------- INTERFAZ --------------------


def seleccionar_archivo():
    ruta = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if ruta:
        archivo_var.set(ruta)
        cargar_hojas(ruta)


def cargar_hojas(ruta):
    try:
        xl = pd.ExcelFile(ruta)
        opciones_hojas.set(xl.sheet_names[0])
        menu_hojas["menu"].delete(0, "end")
        for hoja in xl.sheet_names:
            menu_hojas["menu"].add_command(
                label=hoja, command=lambda h=hoja: opciones_hojas.set(h)
            )
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo cargar el archivo: {e}")


def ejecutar_en_hilo():
    archivo = archivo_var.get()
    hoja = opciones_hojas.get()
    columna = columna_var.get().strip().upper()
    columna_extra = columna_extra_var.get().strip().upper()
    header = bool(tiene_header_var.get())

    if not os.path.isfile(archivo):
        messagebox.showerror("Error", "Selecciona un archivo válido.")
        return
    if not columna.isalpha():
        messagebox.showerror("Error", "Ingresa una letra de columna válida (A-Z).")
        return

    btn_ejecutar.config(state="disabled")
    label_proceso.config(text="Procesando...")

    def proceso():
        try:
            nuevo = procesar_excel_con_pandas(
                archivo, hoja, columna, header, columna_extra if columna_extra else None
            )
            root.after(0, lambda: mostrar_opcion_abrir_archivo(nuevo))
        except Exception as err:
            root.after(0, lambda: messagebox.showerror("Error", str(err)))
        finally:
            root.after(0, lambda: btn_ejecutar.config(state="normal"))
            root.after(0, lambda: label_proceso.config(text="Proceso completado."))

    threading.Thread(target=proceso).start()


def mostrar_opcion_abrir_archivo(archivo_guardado):
    respuesta = messagebox.askyesno(
        "Archivo guardado",
        f"Archivo guardado como:\n{archivo_guardado}\n\n¿Quieres abrirlo?",
    )
    if respuesta:
        abrir_archivo(archivo_guardado)


def abrir_archivo(archivo):
    try:
        if os.name == "nt":
            subprocess.Popen(["start", archivo], shell=True)
        else:
            subprocess.call(["open", archivo])
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo abrir el archivo: {e}")


# -------------------- UI --------------------

root = Tk()
root.title("Duplicador de Grupos en Excel (Optimizado)")
root.geometry("520x350")

archivo_var = StringVar()
columna_var = StringVar(value="B")
columna_extra_var = StringVar(value="H")
opciones_hojas = StringVar()
tiene_header_var = IntVar(value=1)

Label(root, text="Archivo Excel:").pack(pady=5)
Button(root, text="Seleccionar archivo", command=seleccionar_archivo).pack()
Label(root, textvariable=archivo_var, wraplength=450).pack(pady=5)

Label(root, text="Seleccionar hoja:").pack()
menu_hojas = OptionMenu(root, opciones_hojas, "")
menu_hojas.pack()

Label(root, text="Columna principal (ej: A, B, C):").pack(pady=5)
Entry(root, textvariable=columna_var).pack()

Label(root, text="Columna extra para validar (opcional, ej: H):").pack(pady=5)
Entry(root, textvariable=columna_extra_var).pack()

Checkbutton(root, text="Tiene encabezado", variable=tiene_header_var).pack(pady=5)

btn_ejecutar = Button(
    root, text="Ejecutar duplicación", command=ejecutar_en_hilo, bg="lightgreen"
)
btn_ejecutar.pack(pady=10)

label_proceso = Label(root, text="", fg="blue")
label_proceso.pack(pady=10)

root.mainloop()
