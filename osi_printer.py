""" OSI Label Printer V1.7
    Programa para imprimir etiquetas con el número de OSI en una impresora térmica conectada al puerto USB.
    Está configurada para la Nippon NP-3511D y NP-3511D-2 con etiquetas de
    70 o 74 mm de ancho x 40 mm, 49 mm o 50 mm de alto con 3, 4 o 5 mm entre etiquetas.
    Autor: Fernando Agostino
    Modificacion: Chazarreta Matias
    Fecha: 14/12/2024
"""

import logging
from datetime import datetime
from win32 import win32print
from escpos.printer import Dummy
import customtkinter as ctk
from customtkinter import filedialog
import pandas as pd

VERSION = "V.C01"

# Poner False en condiciones normales o True si se quieren imprimir los mensajes para debug
DEBUG = False

# Poner True en condiciones normales o False si no se quiere enviar a la impresora
IMPRIMIR = True

# Establezco el nivel de logging
if DEBUG:
    logging.basicConfig(level=logging.DEBUG)
else:
    logging.basicConfig(level=logging.INFO)

# Defino algunas constantes
ESC = b"\x1b"           # Escape char
GS  = b"\x1d"           # Group Separator char
CTL_LF = b"\n"          # Print and line feed
CTL_FF = b"\f"          # Form feed
CTL_CR = b"\r"          # Carriage return
CTL_HT = b"\t"          # Horizontal tab
CUTTER_OFFSET = 9.5     # Distancia en mm entre el cabezal de impresión y el cutter
POST_CUT_ADV = 2        # Distancia en mm que la impresora avanza el papel luego del corte
MARGIN = 4              # Distancia en mm desde borde de etiqueta a primera línea impresa
AJUSTE = 6.0            # Distancia en mm a avanzar al final de la etiqueta. Estimado con prueba y error para evitar corrimiento

class Np3511d(Dummy):
    """ Subclase para la ipresora Nippon NP-3511D y NP-3511D-2 """

    def set_max_print_speed(self, speed):
        """ Establece la velocidad de impresión en 75, 100, 125, 150 o 200 mm/seg """

        match speed:
            case 200:
                self._raw(GS + b"\x53" + b"\x00") # Comando para setear max. velocidad a 200 mm/s
            case 150:
                self._raw(GS + b"\x53" + b"\x01") # Comando para setear max. velocidad a 150 mm/s
            case 125:
                self._raw(GS + b"\x53" + b"\x02") # Comando para setear max. velocidad a 125 mm/s
            case 100:
                self._raw(GS + b"\x53" + b"\x03") # Comando para setear max. velocidad a 100 mm/s
            case _:
                self._raw(GS + b"\x53" + b"\x04") # Comando para setear max. velocidad a  75 mm/s


    def set_alignment(self, alignment):
        """ Establece la alineación a "LEFT", "CENTER" o "RIGHT" """
        match alignment:
            case "RIGHT":
                self._raw(ESC + b"\x61" + b"\x02") # Comando para alinear a la derecha
            case "CENTER":
                self._raw(ESC + b"\x61" + b"\x01") # Comando para alinear al centro
            case _:
                self._raw(ESC + b"\x61" + b"\x00") # Comando para alinear a la izquierda


    def set_print_density(self, density):
        """ Establece la densidad de impresión entre 65% y 130% """

        if density < 65:
            density = 65
        elif density > 130:
            density = 130

        self._raw(GS + b"\x7E" + bytes([int(density)])) # Comando para setear la densidad de impresión al valor indicado


    def set_enhanced_print_on(self):
        """ Enciende la impresión mejorada """

        self._raw(ESC + b"\x45" + b"\x01") # Comando para setear la impresión mejorada


    def set_enhanced_print_off(self):
        """ Apaga la impresión mejorada """
        self._raw(ESC + b"\x45" + b"\x00") # Comando para resetear la impresión mejorada


    def set_double_strike_on(self):
        """ Enciende la doble pasada """

        self._raw(ESC + b"\x47" + b"\x01") # Comando para setear Double Strike


    def set_double_strike_off(self):
        """ Apaga la doble pasada """

        self._raw(ESC + b"\x47" + b"\x00") # Comando para resetear Double Strike


    def set_double_width_and_height(self):
        """ Enciende doble alto y doble ancho con Font A """

        self._raw(ESC + b"\x21" + b"\x38") # Comando para setear doble ancho y doble alto con Font A


    def set_normal_width_and_height(self):
        """ Setaea alto y ancho normal con Font A """

        self._raw(ESC + b"\x21" + b"\x00") # Comando para setear ancho y alto normal con Font A


    def set_barcode_width(self, width):
        """ Establece el ancho horizontal del código de barras en 2, 3 o 4 """

        match width:
            case 4:
                self._raw(GS + b"\x77" + b"\x04") # Comando para setear el ancho horizontal del barcode en 4
            case 2:
                self._raw(GS + b"\x77" + b"\x02") # Comando para setear el ancho horizontal del barcode en 2
            case _:
                self._raw(GS + b"\x77" + b"\x03") # Comando para setear el ancho horizontal del barcode en 3 (normal)


    def full_cut(self):
        """ Realiza un corte total del papel """

        self._raw(ESC + b"\x69") # Comando para corte total del papel


    def reset(self):
        """ Resetea la impresora """

        self._raw(ESC + b"\x40") # Comando para resetear la impresora


    def set_lf_pitch(self, pitch):
        """ Setea el menor valor de line feed en múltiplos de 1/203 de pulgada (aprox 1/8 mm) """

        if pitch < 0:
            pitch = 0
        elif pitch > 255:
            pitch = 255

        self._raw(ESC + b"\x33" + bytes([int(pitch)])) # Comando para setear el valor del line feed


    def feed_forward_mm(self, distance):
        """ Avanza el papel hacia adelante la cantidad de milímetros indicada """

        if distance < 0:
            return self.feed_backward_mm(-distance)

        multiplos = int(distance / 31.875)
        resto = distance % 31.875

        while multiplos > 0:
            self._raw(ESC + b"\x4A" + bytes([int(31.875 * 8)])) # Comando para avanzar 31.875 milímetros
            multiplos -=1

        self._raw(ESC + b"\x4A" + bytes([int(resto * 8)])) # Comando para avanzar "n" milímetros

        return distance


    def feed_backward_mm(self, distance):
        """ Avanza el papel hacia atrás la cantidad de milímetros indicada """

        if distance < 0:
            return self.feed_forward_mm(-distance)

        multiplos = int(distance / 31.875)
        resto = distance % 31.875

        while multiplos > 0:
            self._raw(ESC + b"\x42" + bytes([int(31.875 * 8)])) # Comando para retroceder 31.875 milímetros
            multiplos -=1

        self._raw(ESC + b"\x42" + bytes([int(resto * 8)])) # Comando para retroceder "n" milímetros

        return -distance


    def set_margins(self, margen_izq, ancho_impresion):
        """ Establece el margen izquierdo y el ancho de impresión en mm """

        self._raw(GS + b"\x57" + bytes([int(margen_izq)]) + bytes([int(ancho_impresion)]))


def print_buffer(buffer):
    """ Imprime el buffer pasado como parámetro en la impresora de Windows """

    # Selecciono la impresora
    my_printer = DEFAULT_PRINTER

    # Impresión de diagnóstico del buffer de impresión en pantalla
    logging.debug(buffer.hex())

    if IMPRIMIR:

        # Imprimo lo generado por la impresora de etiquetas en la impresora de Windows
        win_printer = win32print.OpenPrinter(my_printer)
        try:
            win_print_job = win32print.StartDocPrinter(
                win_printer, 1, ("OSI Label Printing", None, "RAW")
            )
            try:
                win32print.StartPagePrinter(win_printer)

                # Imprimo la etiqueta
                win32print.WritePrinter(win_printer, buffer)

                win32print.EndPagePrinter(win_printer)
                logging.debug("Imprimiendo...")
            finally:
                win32print.EndDocPrinter(win_printer)
        finally:
            # Cierro la impresora de windows
            win32print.ClosePrinter(win_printer)
            win_printer = None
            logging.debug("Etiqueta impresa")


def button_print_callback():
    """ Realiza la impresión de la etiqueta seleccionada """

    # Creo una impresora de tickets Dummy para generar los comandos que luego envío a la impresora real
    p = Np3511d()

    # Reseteo la configuración de la impresora
    p.reset()

    # Establezco el paso del line feed al mínimo
    p.set_lf_pitch(0)

    # Seteo Font A en doble alto y doble ancho
    p.set_double_width_and_height()

    # Seteo la alineación a la izquierda
    p.set_alignment("LEFT")

    # Pruebo algunos seteos para mejorar la calidad de impresión
    # Limito la velocidad de impresión
    p.set_max_print_speed(75)
    # Subo la densidad de impresión
    p.set_print_density(130)
    # Subo la calidad de impresión
    p.set_enhanced_print_on()
    # Seteo impresión de doble pasada
    p.set_double_strike_on()

    # Seteo el margen izquierdo en 6 mm y el ancho de impresión en 66 mm
    p.set_margins(6, 66)

    # Leo el tipo de etiqueta elegida
    tipo_etiqueta = optionmenu_label_type.get()

    # Leo el tamaño de etiqueta elegida
    alto_etiqueta = int(optionmenu_label_size.get())

    # Leo separación entre etiquetas elegida
    gap = int(optionmenu_gap_size.get())

    # Retrocedo desde la línea de corte para posicionar el cabezal en la primera línea de la siguiente etiqueta
    p.feed_backward_mm(CUTTER_OFFSET - gap/2 - MARGIN + POST_CUT_ADV)

    # Reseteo el registro de la posición del cabezal de impresión en la etiqueta
    # contando en mm desde el punto medio de la separación entre etiquetas
    head_pos = MARGIN + gap/2

    match tipo_etiqueta:

        case "Ingreso Equipo":
            # Leo el nro de OSI (debe tener 6 caracteres)
            nro_osi = entrada_osi.get()
            if len(nro_osi) != 6:
                # Cierro la impresora dummy y regreso
                p.clear()
                p.close()
                p = None
                return

            # Calculo la separación en mm entre los bloques de impresión (sub-etiquetas)
            nro_lineas = 5
            nro_bloques = 3
            separacion = int(((alto_etiqueta - 2 * MARGIN - 6 * nro_lineas) / (nro_bloques - 1)) * 8) / 8

            # Imprimo el Nro de OSI dos veces en la misma línea
            p.text(" OSI " + nro_osi + "\t   *" + nro_osi)
            # Ajusto el registro de la posición del cabezal en mm
            head_pos += 6

            # Imprimo el código de barras de 6 mm sin texto y con los caracteres de START y END en CODE39
            # barcode(code, bc, height=64, width=3, pos='BELOW', font='A', align_ct=True, function_type=None, check=True)
            p.barcode("*" + nro_osi + "*", "CODE39", width=3, height=48, pos="OFF", align_ct=False)
            # Ajusto el registro de la posición del cabezal en mm
            head_pos += 6

            # Dejo espacio entre bloques y ajusto el registro de la posición del cabezal en mm
            head_pos += p.feed_forward_mm(separacion)

            # Vuelvo a imprimir el Nro de OSI dos veces en la misma línea
            p.text(" OSI " + nro_osi + "\t   *" + nro_osi)
            # Ajusto el registro de la posición del cabezal en mm
            head_pos += 6

            # Vuelvo a imprimir el código de barras de 6 mm sin texto y con los caracteres de START y END en CODE39
            p.barcode("*" + nro_osi + "*", "CODE39", width=3, height=48, pos="OFF", align_ct=False)
            # Ajusto el registro de la posición del cabezal en mm
            head_pos += 6

            # Dejo espacio entre bloques y ajusto el registro de la posición del cabezal en mm
            head_pos += p.feed_forward_mm(separacion)

            # Vuelvo a imprimir el Nro de OSI dos veces en la misma línea
            p.text(" OSI " + nro_osi + "\t   *" + nro_osi)
            # Ajusto el registro de la posición del cabezal en mm
            head_pos += 6

            logging.debug("Separacion entre bloques = %f", separacion)
            logging.debug("Posicion Cabezal antes del corte = %f", head_pos)

            # Avanzo el papel hasta el punto de corte, dependiendo del alto de etiqueta y su separación
            head_pos += p.feed_forward_mm(alto_etiqueta + gap - head_pos + CUTTER_OFFSET + AJUSTE)

            logging.debug("Posicion Cabezal luego del corte = %f", head_pos)

        case "Ingreso Golden Unit":
            # Leo el nro de OSI (debe tener 6 caracteres)
            nro_osi = entrada_osi.get()
            if len(nro_osi) != 6:
                # Cierro la impresora dummy y regreso
                p.clear()
                p.close()
                p = None
                return

            # Calculo la separación en mm entre los bloques de impresión (sub-etiquetas)
            nro_lineas = 5
            nro_bloques = 3
            separacion = int(((alto_etiqueta - 2 * MARGIN - 6 * nro_lineas) / (nro_bloques - 1)) * 8) / 8

            # Imprimo el Nro de OSI dos veces en la misma línea precedido por GU o _
            p.text("  GU " + nro_osi + "\t   _" + nro_osi)
            # Ajusto el registro de la posición del cabezal en mm
            head_pos += 6

            # Imprimo el código de barras de 6 mm sin texto y con los caracteres de START y END en CODE39
            # barcode(code, bc, height=64, width=3, pos='BELOW', font='A', align_ct=True, function_type=None, check=True)
            p.barcode("*" + nro_osi + "*", "CODE39", height=48, pos="OFF", align_ct=False)
            # Ajusto el registro de la posición del cabezal en mm
            head_pos += 6

            # Dejo espacio entre bloques y ajusto el registro de la posición del cabezal en mm
            head_pos += p.feed_forward_mm(separacion)

            # Vuelvo a imprimir el Nro de OSI dos veces en la misma línea
            p.text("  GU " + nro_osi + "\t   _" + nro_osi)
            # Ajusto el registro de la posición del cabezal en mm
            head_pos += 6

            # Vuelvo a imprimir el código de barras de 6 mm sin texto y con los caracteres de START y END en CODE39
            p.barcode("*" + nro_osi + "*", "CODE39", height=48, pos="OFF", align_ct=False)
            # Ajusto el registro de la posición del cabezal en mm
            head_pos += 6

            # Dejo espacio entre bloques y ajusto el registro de la posición del cabezal en mm
            head_pos += p.feed_forward_mm(separacion)

            # Vuelvo a imprimir el Nro de OSI dos veces en la misma línea
            p.text("  GU " + nro_osi + "\t   _" + nro_osi)
            # Ajusto el registro de la posición del cabezal en mm
            head_pos += 6

            logging.debug("Separacion entre bloques = %f", separacion)
            logging.debug("Posicion Cabezal antes del corte = %f", head_pos)

            # Avanzo el papel hasta el punto de corte, dependiendo del alto de etiqueta y su separación
            head_pos += p.feed_forward_mm(alto_etiqueta + gap - head_pos + CUTTER_OFFSET + AJUSTE)

            logging.debug("Posicion Cabezal luego del corte = %f", head_pos)

        case "Salida Parte":
            # Leo el nro de OSI (debe tener 6 caracteres)
            nro_osi = entrada_osi.get()
            ''' if len(nro_osi) != 6:
                # Cierro la impresora dummy y regreso
                p.clear()
                p.close()
                p = None
                return '''

            # Calculo la separación en mm entre los bloques de impresión (sub-etiquetas)
            nro_lineas = 5
            nro_bloques = 3
            separacion = int(((alto_etiqueta - 2 * MARGIN - 6 * nro_lineas) / (nro_bloques - 1)) * 8) / 8

            # Leo el Claim y el FRU
            nro_claim = entrada_claim.get()[:14]
            nro_fru = entrada_fru.get()[:14]

            # Obtengo la fecha de hoy en formato texto
            hoy = datetime.now().strftime("%d/%m/%Y")

            # Imprimo la fecha de hoy
            p.text("           " + hoy + "\n")
            # Ajusto el registro de la posición del cabezal en mm
            head_pos += 6

            # Dejo espacio entre bloques y ajusto el registro de la posición del cabezal en mm
            head_pos += p.feed_forward_mm(separacion)

            # Imprimo el Nro de FRU
            p.text(" FRU:   " + nro_fru + "\n")
            # Ajusto el registro de la posición del cabezal en mm
            head_pos += 6

            # Imprimo el Nro de Claim
            p.text(" CLAIM: " + nro_claim + "\n")
            # Ajusto el registro de la posición del cabezal en mm
            head_pos += 6

            # Dejo espacio entre bloques y ajusto el registro de la posición del cabezal en mm
            head_pos += p.feed_forward_mm(separacion)

            # Imprimo el Nro de OSI
            p.text(" OSI:   " + nro_osi + "\n")
            # Ajusto el registro de la posición del cabezal en mm
            head_pos += 6

            # Imprimo el código de barras de 6 mm sin texto y con los caracteres de START y END en CODE39
            p.barcode("*" + nro_osi + "*", "CODE39", width=4, height=48, pos="OFF", align_ct=False)
            # Ajusto el registro de la posición del cabezal en mm
            head_pos += 6

            logging.debug("Separacion entre bloques = %f", separacion)
            logging.debug("Posicion Cabezal antes del corte = %f", head_pos)

            # Avanzo el papel hasta el punto de corte, dependiendo del alto de etiqueta y su separación
            head_pos += p.feed_forward_mm(alto_etiqueta + gap - head_pos + CUTTER_OFFSET)

            logging.debug("Posicion Cabezal luego del corte = %f", head_pos)

        case "Libre":
            # Calculo la separación en mm entre los bloques de impresión (sub-etiquetas)
            nro_lineas = 5
            nro_bloques = 5
            separacion = int(((alto_etiqueta - 2 * MARGIN - 6 * nro_lineas) / (nro_bloques - 1)) * 8) / 8

            lineas = []
            # Leo los renglones a imprimir (limitados a 20 caracteres por línea)
            lineas.append(entrada_osi.get()[:20])
            lineas.append(entrada_claim.get()[:20])
            lineas.append(entrada_fru.get()[:20])
            lineas.append(entrada_4ta_linea.get()[:20])
            lineas.append(entrada_5ta_linea.get()[:20])

            for linea in lineas:
                # Imprimo la línea. Si la línea está vacía, immprimo un espacio
                if len(linea) > 0:
                    p.text(linea + "\n")
                else:
                    p.text(" \n")

                # Ajusto el registro de la posición del cabezal en mm
                head_pos += 6

                # Dejo espacio entre bloques y ajusto el registro de la posición del cabezal en mm
                head_pos += p.feed_forward_mm(separacion)

            logging.debug("Separacion entre bloques = %f", separacion)
            logging.debug("Posicion Cabezal antes del corte = %f", head_pos)

            # Avanzo el papel hasta el punto de corte, dependiendo del alto de etiqueta y su separación
            head_pos += p.feed_forward_mm(alto_etiqueta + gap - head_pos + CUTTER_OFFSET)

            logging.debug("Posicion Cabezal luego del corte = %f", head_pos)
        

    # Corto la etiqueta
    p.full_cut()

    # Imprimo el buffer de impresión en la impresora de etiquetas
    print_buffer(p.output)

    # Cierro la impresora dummy
    p.clear()
    p.close()
    p = None

    # Me fijo si tengo que borrar los campos de entrada
    if checkbox_borrado.get():
        # Reseteo los campos de entrada para el tipo de etiqueta en uso
        label_type_callback(tipo_etiqueta)
        # Vuelvo a poner el foco en el campo de ingreso del nro de OSI
        entrada_osi.focus_set()


def button_forward_callback():
    """ Avanza el papel 0.5 mm """

    # Creo una impresora de tickets Dummy para generar los comandos que luego envío a la impresora real
    p = Np3511d()

    # Avanzo el papel 0.5 mm
    p.feed_forward_mm(0.5)

    # Imprimo el buffer de impresión en la impresora de etiquetas
    print_buffer(p.output)

    # Cierro la impresora dummy
    p.clear()
    p.close()
    p = None


def button_backward_callback():
    """ Retrocede el papel 0.5 mm """

    # Creo una impresora de tickets Dummy para generar los comandos que luego envío a la impresora real
    p = Np3511d()

    # Retrocedo el papel 1 mm
    p.feed_backward_mm(0.5)

    # Imprimo el buffer de impresión en la impresora de etiquetas
    print_buffer(p.output)

    # Cierro la impresora dummy
    p.clear()
    p.close()
    p = None

def print_excel():
    filepath = filedialog.askopenfile()
    df = pd.read_excel(filepath.name)
    """ Realiza la impresión de la etiqueta seleccionada """

    # Creo una impresora de tickets Dummy para generar los comandos que luego envío a la impresora real
    p = Np3511d()

    # Reseteo la configuración de la impresora
    p.reset()

    # Establezco el paso del line feed al mínimo
    p.set_lf_pitch(0)

    # Seteo Font A en doble alto y doble ancho
    p.set_double_width_and_height()

    # Seteo la alineación a la izquierda
    p.set_alignment("LEFT")

    # Pruebo algunos seteos para mejorar la calidad de impresión
    # Limito la velocidad de impresión
    p.set_max_print_speed(75)
    # Subo la densidad de impresión
    p.set_print_density(130)
    # Subo la calidad de impresión
    p.set_enhanced_print_on()
    # Seteo impresión de doble pasada
    p.set_double_strike_on()

    # Seteo el margen izquierdo en 6 mm y el ancho de impresión en 66 mm
    p.set_margins(6, 66)

    # Leo el tamaño de etiqueta elegida
    alto_etiqueta = int(optionmenu_label_size.get())

    nro_lineas = 5
    nro_bloques = 3

    # Leo separación entre etiquetas elegida
    gap = int(optionmenu_gap_size.get())

    # Retrocedo desde la línea de corte para posicionar el cabezal en la primera línea de la siguiente etiqueta
    p.feed_backward_mm(CUTTER_OFFSET - gap/2 - MARGIN + POST_CUT_ADV)

    # Reseteo el registro de la posición del cabezal de impresión en la etiqueta
    # contando en mm desde el punto medio de la separación entre etiquetas
    head_pos = MARGIN + gap/2
    i=0
    for value in df.values:
        i+=1
        
        separacion = int(((alto_etiqueta - 2 * MARGIN - 6 * nro_lineas) / (nro_bloques - 1)) * 8) / 8
        nro_osi = str(value[0])
        nro_claim = str(value[1])
        nro_fru = str(value[2])
        if i % 5 == 0:
            button_forward_callback()

        # Obtengo la fecha de hoy en formato texto
        hoy = datetime.now().strftime("%d/%m/%Y")

        # Imprimo la fecha de hoy
        p.text("           " + hoy + "\n")
        # Ajusto el registro de la posición del cabezal en mm
        head_pos += 6

        # Dejo espacio entre bloques y ajusto el registro de la posición del cabezal en mm
        head_pos += p.feed_forward_mm(separacion)

        # Imprimo el Nro de FRU
        p.text(" FRU:   " + nro_fru + "\n")
        # Ajusto el registro de la posición del cabezal en mm
        head_pos += 6

        # Imprimo el Nro de Claim
        p.text(" CLAIM: " + nro_claim + "\n")
        # Ajusto el registro de la posición del cabezal en mm
        head_pos += 6

        # Dejo espacio entre bloques y ajusto el registro de la posición del cabezal en mm
        head_pos += p.feed_forward_mm(separacion)

        # Imprimo el Nro de OSI
        p.text(" OSI:   " + nro_osi + "\n")
        # Ajusto el registro de la posición del cabezal en mm
        head_pos += 6

        # Imprimo el código de barras de 6 mm sin texto y con los caracteres de START y END en CODE39
        p.barcode("*" + nro_osi + "*", "CODE39", width=4, height=48, pos="OFF", align_ct=False)
        # Ajusto el registro de la posición del cabezal en mm
        head_pos += 6

        logging.debug("Separacion entre bloques = %f", separacion)
        logging.debug("Posicion Cabezal antes del corte = %f", head_pos)

        # Avanzo el papel hasta el punto de corte, dependiendo del alto de etiqueta y su separación
        head_pos += p.feed_forward_mm(alto_etiqueta + gap - head_pos + CUTTER_OFFSET)

        logging.debug("Posicion Cabezal luego del corte = %f", head_pos)

        # Corto la etiqueta
        p.full_cut()

        # Imprimo el buffer de impresión en la impresora de etiquetas
        print_buffer(p.output)

    # Cierro la impresora dummy
    p.clear()
    p.close()
    p = None



def limpiar_entradas():
    """ Borra todos los campos de entrada """
    entrada_osi.delete(0, ctk.END)
    entrada_claim.delete(0, ctk.END)
    entrada_fru.delete(0, ctk.END)
    entrada_4ta_linea.delete(0, ctk.END)
    entrada_5ta_linea.delete(0, ctk.END)


def label_type_callback(label_type):
    """ Setea los campos de entrada según el tipo de etiqueta """

    # Borro todos los campos de entrada
    limpiar_entradas()

    # Configuro los campos de entrada según el tipo de etiqueta
    match label_type:
        case "Ingreso Equipo" | "Ingreso Golden Unit":
            # Seteo solo el campo de Nro de OSI
            entrada_osi.configure(placeholder_text="Nro de OSI", state=ctk.NORMAL)
            entrada_claim.configure(placeholder_text="", state=ctk.NORMAL)
            entrada_claim.configure(state=ctk.DISABLED)
            entrada_fru.configure(placeholder_text="", state=ctk.NORMAL)
            entrada_fru.configure(state=ctk.DISABLED)
            entrada_4ta_linea.configure(placeholder_text="", state=ctk.NORMAL)
            entrada_4ta_linea.configure(state=ctk.DISABLED)
            entrada_5ta_linea.configure(placeholder_text="", state=ctk.NORMAL)
            entrada_5ta_linea.configure(state=ctk.DISABLED)

        case "Salida Parte":
            # Seteo solo los campos de Nro de OSI, Claim y FRU
            entrada_osi.configure(placeholder_text="Nro de OSI", state=ctk.NORMAL)
            entrada_claim.configure(placeholder_text="Nro de CLAIM",state=ctk.NORMAL)
            entrada_fru.configure(placeholder_text="Nro de FRU", state=ctk.NORMAL)
            entrada_4ta_linea.configure(placeholder_text="", state=ctk.NORMAL)
            entrada_4ta_linea.configure(state=ctk.DISABLED)
            entrada_5ta_linea.configure(placeholder_text="", state=ctk.NORMAL)
            entrada_5ta_linea.configure(state=ctk.DISABLED)

        case "Libre":
            # Seteo los cinco campos como líneas 1 a 5
            entrada_osi.configure(placeholder_text="1ra Línea", state=ctk.NORMAL)
            entrada_claim.configure(placeholder_text="2da Línea", state=ctk.NORMAL)
            entrada_fru.configure(placeholder_text="3ra Línea", state=ctk.NORMAL)
            entrada_4ta_linea.configure(placeholder_text="4ta Línea", state=ctk.NORMAL)
            entrada_5ta_linea.configure(placeholder_text="5ta Línea", state=ctk.NORMAL)

        case "Excel":
            print("")

    # Saco el foco de los campos de entrada
    app.focus_set()



#####################################################################################################################
# COMIENZO DEL PROGRAMA
#####################################################################################################################

#if __name__ == "__main__":

# Obtengo la lista de impresoras disponibles en la PC
# printers = win32print.EnumPrinters(2)
# printer_list = []
# for x, printer in enumerate(printers):
#    printer_list.append(printer[2])

# Seteo mi impresora como el nombre por defecto de la Nippon
DEFAULT_PRINTER = "NPI Integration Driver"

# Armo la lista de tipos de etiquetas disponibles
label_type_list = ["Salida Parte","Ingreso Equipo", "Ingreso Golden Unit", "Libre", "Excel"]
DEFAULT_LABEL_TYPE = label_type_list[0]

# Armo la lista de tamaños de etiquetas disponibles (altura en mm)
label_size_list = ["50","40", "49"]
DEFAULT_LABEL_SIZE = label_size_list[0]

# Armo la lista de separación entre etiquetas disponibles (espacio en mm)
gap_size_list = [ "5","6", "4"]
DEFAULT_GAP_SIZE = gap_size_list[0]

# Configura la ventana principal
ctk.set_appearance_mode("dark")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"
app = ctk.CTk()
app.geometry("290x630")
app.title("OSI Label Printer " + VERSION)

# Ingreso de Nro de OSI
entrada_osi = ctk.CTkEntry(master=app)
entrada_osi.pack(pady=(15, 5))

# Ingreso de Nro de Claim
entrada_claim = ctk.CTkEntry(master=app)
entrada_claim.pack(pady=5)

# Ingreso de Nro de FRU
entrada_fru = ctk.CTkEntry(master=app)
entrada_fru.pack(pady=5)

# Ingreso de 4ta línea
entrada_4ta_linea = ctk.CTkEntry(master=app)
entrada_4ta_linea.pack(pady=5)

# Ingreso de 5ta línea
entrada_5ta_linea = ctk.CTkEntry(master=app)
entrada_5ta_linea.pack(pady=5)

# Botón para Imprimir la etiqueta
boton_imprimir = ctk.CTkButton(master=app, text="IMPRIMIR", command=button_print_callback)
boton_imprimir.pack(pady=25)

# Menú de selección de Impresora
# optionmenu_printer = ctk.CTkOptionMenu(master=app, values=printer_list)
# optionmenu_printer.pack(pady=10, padx=10)
# optionmenu_printer.set(DEFAULT_PRINTER)

# Selecciono el tipo de etiqueta a imprimir
title_label_type = ctk.CTkLabel(master=app, justify=ctk.LEFT, text="Tipo de Etiqueta")
title_label_type.pack(pady=1, anchor="s")
optionmenu_label_type = ctk.CTkOptionMenu(master=app, values=label_type_list, command=label_type_callback)
optionmenu_label_type.pack(pady=1, anchor="n")
optionmenu_label_type.set(DEFAULT_LABEL_TYPE)

# Selecciono la altura de la etiqueta
title_label_size = ctk.CTkLabel(master=app, justify=ctk.LEFT, text="Altura Etiqueta (mm)")
title_label_size.pack(pady=1, anchor="s")
optionmenu_label_size = ctk.CTkOptionMenu(master=app, values=label_size_list)
optionmenu_label_size.pack(pady=1, anchor="n")
optionmenu_label_size.set(DEFAULT_LABEL_SIZE)

# Selecciono la espacio entre etiquetas (gap)
title_gap_size = ctk.CTkLabel(master=app, justify=ctk.LEFT, text="Separación Etiqueta (mm)")
title_gap_size.pack(pady=1, anchor="s")
optionmenu_gap_size = ctk.CTkOptionMenu(master=app, values=gap_size_list)
optionmenu_gap_size.pack(pady=1, anchor="n")
optionmenu_gap_size.set(DEFAULT_GAP_SIZE)

# Agrego checkbox para ver si borro o no los campos entre etiqueta y etiqueta
checkbox_borrado = ctk.CTkCheckBox(master=app, text="Borrar Entradas")
checkbox_borrado.pack(pady=15)

# Botón para avanzar el papel
boton_avanzar = ctk.CTkButton(master=app, text="AVANZAR", command=button_forward_callback)
boton_avanzar.pack(pady=5)

# Botón para retroceder el papel
boton_retroceder = ctk.CTkButton(master=app, text="RETROCEDER", command=button_backward_callback)
boton_retroceder.pack(pady=5)

# Boton para abrir el excel

boton_excel = ctk.CTkButton(master=app,text="EXCEL",command=print_excel)
boton_excel.pack(pady=5)

# Seteo los campos de entrada para el tipo de etiqueta por defecto
label_type_callback(DEFAULT_LABEL_TYPE)

# Lazo principal
app.mainloop()
