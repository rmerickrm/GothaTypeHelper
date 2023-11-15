#biblioteca docx es necesario instalar mediante pip install
import docx
import tkinter as tk    #Graficos
from tkinter import ttk #esto es para la barra de progreso y objetos estilizados
import tkinter.scrolledtext as scrolledtext #textarea mejorada
import pyperclip        #portapapeles
import argparse         #recepción de parámetros


# Globales
i = 0
total_parrafos = 0
texto_listado = list()


#-------------------------------------------------------------------------------------------------------
# Parte Lógica
#-------------------------------------------------------------------------------------------------------


# Necesito agregar control de extención de documento y/o control de archivos no válidos
def leer_documento(mi_documento):
    """
    Recibe un documento por parámetro y llena una lista con todos sus párrafos no vacíos.
    :param mi_documento: Nombre del documento a abrir..
    """
    print(mi_documento)
    global texto_listado
    global total_parrafos
    documento_recibido = docx.Document(mi_documento)
    for paragraph in documento_recibido.paragraphs:
        texto_paragraph = paragraph.text
        if len(texto_paragraph) != 0:
            texto_listado.append(texto_paragraph)
            total_parrafos += 1

def no_documento():
    """
    Llena la lista de párrafos con el mensaje 'Sin información que mostrar' e indica que sólo tiene un párrafo, se
    utiliza cuando no se ingresa ningún parámetro.
    """
    global texto_listado
    global total_parrafos
    texto_listado.append("Sin Información que mostrar.")
    total_parrafos = 1


def progreso(mi_posicion):
    """
    Asigna la posición en la barra de progreso basándose en el número recibido y la cantidad de párrafos totales del
    documento.
    :param mi_posicion: Posición actual del párrafo.
    """
    mi_posicion
    porcentaje = int(mi_posicion * 100 / total_parrafos)
    barra_progreso['value'] = porcentaje


def leer_elementos():
    """
    Lee el siguiente elemento o párrafo en el documento y lo asigna al cuadro de texto y portapapeles.
    """
    global i
    largo_lista = len(texto_listado)
    # Mejor prevenir aunque nunca debería ser cierto
    if i < largo_lista:
        texto = cuadro_texto.get("1.0", "end")
        cuadro_texto.delete("1.0", "end")
        if mayusculas.get() == 1:
            texto_listado[i] = texto_listado[i].upper()
        largo_texto = len(texto_listado[i])
        cuadro_texto.insert("1.0", texto_listado[i])
        if largo_texto > 1 and texto_listado[i][0] == "*" and texto_listado[i][largo_texto - 1] == "*":
            texto = texto_listado[i][1:largo_texto - 1]
            cuadro_texto.configure(foreground="#600000")
        elif largo_texto > 1 and texto_listado[i][0] == "*":
            texto = texto_listado[i][1:largo_texto]
            cuadro_texto.configure(foreground="#600000")
        else:
            texto = texto_listado[i]
            cuadro_texto.configure(foreground="black")
        pyperclip.copy(texto)
        mostrar_posicion(i)
        # Posición del siguiente texto
        if i + 1 < largo_lista:
            i = i + 1
            progreso(i)


def anterior():
    """
    Decrementa el índice de la posición del párrafo para que el párrafo anterior sea leído.
    """
    global i
    # Posición del texto actual es i-1 se resta dos porque al llamar a leer se avanza siempre 1
    i = i - 2
    # No queremos underflow
    if i < 0:
        i = 0
    leer_elementos()


def copiarPortapapeles():
    """
    Copia a portapapeles el texto exácto que se encuentra en el cuadro de texto.
    """
    global i
    if i > 0:
        pyperclip.copy(texto_listado[i - 1])
    else:
        pyperclip.copy("")


def mostrar_posicion(posicion_actual):
    """
    Borra lo anteriormente escrito y asigna mi posición de párrafo actual en el elemento gráfico correspondiente.
    :param posicion_actual: Mi número de párrafo actual.
    """
    parrafo_actual.delete(0, "end")
    parrafo_actual.insert(0, posicion_actual + 1)


def ir():
    """
    Desplaza el índice y coloca el texto correspondiente a la posicion indicada en el elemento gráfico parrafo_actual
    además si el número indicado se sale de los límites aplica color rojo en el texto, de lo contrario aplica color negro.
    """
    global i
    numero_parrafo = int(parrafo_actual.get())
    if 0 < numero_parrafo <= total_parrafos:
        parrafo_actual.configure(foreground="black")
        i = numero_parrafo - 1
        leer_elementos()
    else:
        parrafo_actual.configure(foreground="red")


def es_numero(char):
    """
    Indica si se recibió un número.
    :param char: Un caracter cualquiera.
    :return: True si el caracter corresponde a un número.
    """
    try:
        int(char)
        return True
    except ValueError:
        return False


def validar_entrada(char):
    """
    Valida si el parámetro recibido es un número o si no se recibió nada.
    :param char: Un caracter.
    :return: True si corresponde a un número o nada.
    """
    if es_numero(char) or char == "":
        return True
    else:
        return False


#-------------------------------------------------------------------------------------------------------
# Recepción de parámetros
#-------------------------------------------------------------------------------------------------------

# Crea un objeto ArgumentParser
parser = argparse.ArgumentParser(description='Procesa un archivo.')

# Agrega un argumento "archivo" para el nombre del archivo
parser.add_argument('archivo', type=str, nargs='?', default=None, help='Nombre del archivo a procesar')

# Analiza los argumentos de la línea de comandos
args = parser.parse_args()

# Obtiene el nombre del archivo ingresado por el usuario
if args.archivo:
    print("hay documento")
    leer_documento(args.archivo)
else:
    print("NOO hay documento")
    no_documento()

#-------------------------------------------------------------------------------------------------------
# Parte Gráfica
#-------------------------------------------------------------------------------------------------------


# Crear una nueva ventana
ventana = tk.Tk()

# Establecer el título de la ventana
ventana.title("Traducción")

# Establecer el tamaño de la ventana
ancho_ventana = ventana.winfo_screenwidth()
alto_ventana = ventana.winfo_screenheight()
alto_ventana = int(alto_ventana - alto_ventana / 10)
altura= str(alto_ventana)
ventana.geometry("150x" + altura + "+{}+{}".format(ancho_ventana - 150, 0))
# Ancho y alto de la ventana porque alteré la geometría anteriormente
ancho_ventana = ventana.winfo_screenwidth()
alto_ventana = ventana.winfo_screenheight()

# Varialble para el checkbox Mayúsculas
mayusculas = tk.IntVar()
# Variable para el número de párrafo
parrafo = tk.IntVar()
parrafo.set(0)


style = ttk.Style()
style.theme_use('clam')
style.configure('my.TButton', font=('Arial', 14), hoverbackground='blue')
style.configure('my.TCheckbutton', font=('Arial', 10), hoverbackground='blue')
style.configure("blue.Horizontal.TProgressbar",  width=8, background='blue')

# Secciones
frame_superior = tk.Frame(ventana)
frame_medio = tk.Frame(ventana)
frame_inferior = tk.Frame(ventana)

# Creación de elementos gráficos
chekear_mayusculas = ttk.Checkbutton(frame_superior, text="Mayúsculas", variable=mayusculas, style='my.TCheckbutton')
parrafo_actual = tk.Entry(frame_superior, textvariable=parrafo, width=3, font=('Arial', 19))
boton_ir = ttk.Button(frame_superior, text="IR", command=ir, width=6)
boton_siguiente = ttk.Button(frame_medio, text="Siguiente", width=8, command=leer_elementos, style='my.TButton')
boton_anterior = ttk.Button(frame_medio, text="Anterior", width=8, command=anterior, style='my.TButton')
cuadro_texto = scrolledtext.ScrolledText(frame_medio, wrap=tk.WORD, height=10)
barra_progreso = ttk.Progressbar(frame_inferior, orient='horizontal', length=150, mode='determinate', style='blue'
    '.Horizontal.TProgressbar')
boton_copiar = ttk.Button(frame_inferior, text="Volver a copiar", command=copiarPortapapeles, style='my.TButton')

# Puesta a punto de algunos elementos
cuadro_texto.configure(background='#427146', font=('Comic Sans MS', 14))
barra_progreso['value'] = 0
frame_superior.configure(height=60)
# Validación del texto ingresado para solo permitir números
parrafo_actual.config(validate="key", validatecommand=(parrafo_actual.register(validar_entrada), '%S'))

# Colocación de todos los elementos en la ventana.
frame_superior.pack(fill=tk.BOTH)
chekear_mayusculas.place(relx=0.0, rely=0.0)
parrafo_actual.place(in_=chekear_mayusculas, relx=0.0, rely=1.1)
boton_ir.place(in_=parrafo_actual, relx=1.1, rely=-0.0)

frame_medio.pack(fill=tk.BOTH, expand=True)
boton_anterior.pack(fill=tk.BOTH)
boton_siguiente.pack(fill=tk.BOTH)
cuadro_texto.pack(fill=tk.BOTH, expand=True)

frame_inferior.pack()
barra_progreso.pack()
boton_copiar.pack()

# Iniciar el ciclo de eventos de la ventana
ventana.mainloop()
