from abc import ABC, abstractmethod
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import datetime
import random
import os


class Jugador:

    def __init__(self, nombre, edad):
        self.__nombre = nombre
        self.__edad = edad

    @property
    def nombre(self):
        return self.__nombre

    @nombre.setter
    def nombre(self, NValor):
        self.__nombre = NValor

    @property
    def edad(self):
        return self.__edad

    @edad.setter
    def edad(self, NValor):
        if NValor > 9:
            self.__edad = NValor


class Partida:

    def __init__(self):
        self.__posicion = 1
        self.__turno = 0
        self.__pierdeT = 0
        self.__resultado = ""

    @property
    def posicion(self):
        return self.__posicion

    @posicion.setter
    def posicion(self, NValor):

        if NValor < 1:
            self.__posicion = 1

        elif NValor > 30:
            self.__posicion = 30

        else:
            self.__posicion = NValor

    @property
    def turno(self):
        return self.__turno

    def sumatoriaT(self):
        self.__turno = self.__turno + 1

    @property
    def pierdeT(self):
        return self.__pierdeT

    @pierdeT.setter
    def pierdeT(self, NValor):
        self.__pierdeT = NValor

    @property
    def resultado(self):
        return self.__resultado

    @resultado.setter
    def resultado(self, NValor):
        self.__resultado = NValor


class Evento(ABC):

    def __init__(self, nombre):
        self.__nombre = nombre

    @property
    def nombre(self):
        return self.__nombre

    @abstractmethod
    def ejecutarEvent(self, partida):
        pass


class FugaAgua(Evento):

    def __init__(self):
        super().__init__("Fuga de agua")

    def ejecutarEvent(self, partida):

        partida.posicion = partida.posicion - 2

        print("FUGA DE AGUA")
        print("Retrocedes 2 casillas")


class CorteElectrico(Evento):

    def __init__(self):
        super().__init__("Corte electrico")

    def ejecutarEvent(self, partida):

        partida.pierdeT = 1

        print("CORTE ELECTRICO")
        print("Pierdes 1 turno")


class TanqueOxigeno(Evento):

    def __init__(self):
        super().__init__("Tanque de oxigeno")

    def ejecutarEvent(self, partida):

        partida.posicion = partida.posicion + 4

        print("TANQUE DE OXIGENO")
        print("Avanzas 4 casillas")


class CriaturaMarino(Evento):

    def __init__(self):
        super().__init__("Criatura marina")

    def ejecutarEvent(self, partida):

        dado = random.randint(1, 6)

        print("CRIATURA MARINA")
        print("Numero del dado:", dado)

        if dado % 2 == 0:

            partida.posicion = 1

            print("Retrocedes al comienzo del juego")

        else:

            partida.resultado = "Derrota-Criatura Marina"

            print("La criatura te derroto")


class Estadisticas:

    def __init__(self):
        self.__historial = []
        self.__archivo_excel = "historial_partidas.xlsx"
        self.cargarDesdeExcel()

    @property
    def historial(self):
        return self.__historial

    def cargarDesdeExcel(self):
        if os.path.exists(self.__archivo_excel):
            try:
                df = pd.read_excel(self.__archivo_excel)
                self.__historial = df.to_dict('records')
                print(f"Historial cargado desde {self.__archivo_excel} ({len(self.__historial)} partidas)")
            except Exception as e:
                print(f"Error al cargar el archivo: {e}")
                self.__historial = []
        else:
            print("No se encontró archivo de historial. Se iniciará uno nuevo.")

    def guardarEnExcel(self):
        try:
            df = pd.DataFrame(self.__historial)
            df.to_excel(self.__archivo_excel, index=False)
            df.to_csv("historial_partidas.csv", index=False)
            print(f"Historial guardado en {self.__archivo_excel}")
        except Exception as e:
            print(f"Error al guardar el archivo: {e}")

    def guardarPartida(self, jugador, partida):

        fecha_hora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        numeroPartida = len(self.__historial) + 1

        self.__historial.append(
            {
                "Jugador": jugador.nombre,
                "Edad": jugador.edad,
                "Resultado": partida.resultado,
                "Turnos": partida.turno,
                "Fecha y hora": fecha_hora,
                "Numero": numeroPartida
            }
        )

        self.guardarEnExcel()

    def mostrarDatos(self):

        if len(self.__historial) == 0:
            print("No hay partidas registradas")
            return

        victorias = 0
        derrotas = 0
        abandonos = 0
        sumaTurnos = 0

        for partida in self.__historial:

            sumaTurnos = sumaTurnos + partida["Turnos"]

            if partida["Resultado"] == "Victoria":

                victorias = victorias + 1

            elif partida["Resultado"] == "Abandono":

                abandonos = abandonos + 1

            else:

                derrotas = derrotas + 1

        promedioTurnos = sumaTurnos / len(self.__historial)

        print("=========================")
        print("======ESTADISTICAS======")
        print("Partidas:", len(self.__historial))
        print("Victorias:", victorias)
        print("Derrotas:", derrotas)
        print("Partidas abandonadas:", abandonos)
        print("Promedio de turnos:", round(promedioTurnos, 2))
        print("=========================")

        print("ULTIMAS PARTIDAS")

        for partida in self.__historial[-3:]:

            print("N de partida:", partida["Numero"])
            print("Jugador:", partida["Jugador"])
            print("Resultado:", partida["Resultado"])
            print("=========================")

    def resultadosPartida(self):

        if len(self.__historial) == 0:
            return

        partida = self.__historial[-1]

        print("=====RESUMEN DE PARTIDA=====")
        print("Jugador:", partida["Jugador"])
        print("Edad:", partida["Edad"])
        print("Resultado:", partida["Resultado"])
        print("Turnos:", partida["Turnos"])
        print("Fecha y hora:", partida["Fecha y hora"])
        print("============================")


class Graficos:

    def __init__(self,ArchivoExcel = "historial_partidas.xlsx"):
        self.__ArchivoExcel = "historial_partidas.xlsx"

    def ReporteGraficos(self):
        try:
            DataframeActual = pd.read_excel(self.__ArchivoExcel)
        except FileNotFoundError:
            print("\n\t        Advertencia")
            print("\t--------------------------------------")
            print("\t|   No se encontró el archivo Excel  |")
            print("\t|        historial_partidas.xlsx     |")
            print("\t--------------------------------------")
            return
            
        if DataframeActual.empty:
            print("\n\t        Advertencia")
            print("\t--------------------------------------")
            print("\t| No hay datos suficientes en Excel  |")
            print("\t|        para generar gráficos       |")
            print("\t--------------------------------------")
            return

        TurnosPorJugador = DataframeActual.groupby("Jugador")["Turnos"].mean()
        ResultadosDistribucion = DataframeActual["Resultado"].value_counts()

        Fig, Ejes = plt.subplots(1, 2, figsize=(14, 6))
        Fig.suptitle("Dashboard Misión Submarina - Gráficos Estándar", fontsize=16, fontweight="bold")

        Ejes[0].pie(ResultadosDistribucion.values, labels=ResultadosDistribucion.index, autopct="%1.1f%%", colors=["green", "red", "gold"], startangle=90, wedgeprops={"edgecolor": "black"})
        Ejes[0].set_title("Distribución de Resultados")

        Ejes[1].bar(TurnosPorJugador.index, TurnosPorJugador.values, color="royalblue", edgecolor="black")
        Ejes[1].set_title("Promedio de Turnos por Jugador")
        Ejes[1].set_xlabel("Jugadores")
        Ejes[1].set_ylabel("Turnos Promedio")
        Ejes[1].grid(axis="y", linestyle="--", alpha=0.7)

        plt.tight_layout()
        plt.show()

    def ReporteGraficosAvanzados(self):
        try:
            DataframeActual = pd.read_excel(self.__ArchivoExcel)
        except FileNotFoundError:
            print("\n\t        Advertencia")
            print("\t--------------------------------------")
            print("\t|   No se encontró el archivo Excel  |")
            print("\t|        historial_partidas.xlsx     |")
            print("\t--------------------------------------")
            return
            
        if DataframeActual.empty:
            print("\n\t        Advertencia")
            print("\t--------------------------------------")
            print("\t| No hay datos suficientes en Excel  |")
            print("\t|      para analisis avanzados       |")
            print("\t--------------------------------------")
            return

        DataframeActual["IndicePartida"] = np.arange(1, len(DataframeActual) + 1)
        AcumuladoTurnos = DataframeActual["Turnos"].cumsum()

        Fig, Ejes = plt.subplots(1, 2, figsize=(14, 6))
        Fig.suptitle("Dashboard Misión Submarina - Gráficos Avanzados", fontsize=16, fontweight="bold")

        Ejes[0].plot(DataframeActual["IndicePartida"], DataframeActual["Turnos"], marker="o", linestyle="-", color="crimson", linewidth=2)
        Ejes[0].set_title("Turnos Consumidos por Partida")
        Ejes[0].set_xlabel("Número de Partida")
        Ejes[0].set_ylabel("Turnos")
        Ejes[0].grid(True, linestyle=":")

        Ejes[1].fill_between(DataframeActual["IndicePartida"], AcumuladoTurnos, color="teal", alpha=0.4, label="Turnos Acumulados")
        Ejes[1].plot(DataframeActual["IndicePartida"], AcumuladoTurnos, color="teal", linewidth=2)
        Ejes[1].set_title("Acumulado de Turnos")
        Ejes[1].set_xlabel("Número de Partida")
        Ejes[1].set_ylabel("Turnos Acumulados")
        Ejes[1].legend(loc="upper left")
        Ejes[1].grid(True, linestyle=":")

        plt.tight_layout()
        plt.show()


class Juego:

    def __init__(self):

        self.jugador = ""
        self.partida = ""
        self.estadisticas = Estadisticas()
        self.graficos = Graficos()

    def mostrarReglas(self):

        print("=====================================")
        print("============= REGLAS ===============")
        print("Llegar a la casilla 30 para ganar")
        print("Maximo 15 turnos")
        print("Casillas | Obstaculos")
        print("5 - 10   : Fuga de Agua")
        print("11 - 18  : Corte Electrico")
        print("19 - 24  : Tanque de Oxigeno")
        print("25 - 29  : Criatura Marina")
        print("=====================================")
        input("Presione ENTER para continuar...")
        
    def registrarJugador(self):

        nombre = input("Nombre: ")

        while True:

            try:
                edad = int(input("Edad: "))
                if edad > 9:
                    break
                print("Edad minima: 10")

            except ValueError:

                print("Dato no valido")

        self.jugador = Jugador(nombre, edad)

        print("Jugador registrado")

    def nuevoJuego(self):

        if self.jugador == "":

            print("Primero debes registrar un jugador")
            return

        self.partida = Partida()

        print("Nueva partida creada")

    def lanzarDados(self):

        return random.randint(1, 6)

    def obtenerEvento(self, pos):

        if pos >= 5 and pos <= 10:

            return FugaAgua()

        elif pos >= 11 and pos <= 18:

            return CorteElectrico()

        elif pos >= 19 and pos <= 24:

            return TanqueOxigeno()

        elif pos >= 25 and pos <= 29:

            return CriaturaMarino()

        return ""

    def iniciarJuego(self):

        if self.partida == "":

            print("Primero debes crear una partida")
            return

        while self.partida.posicion < 30 and self.partida.turno < 15:

            if self.partida.pierdeT == 1:

                print("Pierdes este turno")

                self.partida.pierdeT = 0

                self.partida.sumatoriaT()

                continue

            print("----------------------")
            print("Posicion:", self.partida.posicion)
            print("Oxigeno:", 15 - self.partida.turno)

            input("Presione ENTER para lanzar el dado")

            dado = self.lanzarDados()

            print("Dado:", dado)

            self.partida.posicion = self.partida.posicion + dado

            evento = self.obtenerEvento(self.partida.posicion)

            if evento != "":

                evento.ejecutarEvent(self.partida)

            if self.partida.resultado != "":
                break

            self.partida.sumatoriaT()

            if self.partida.posicion >= 30:

                self.partida.resultado = "Victoria"

                break

            opcion = input("Desea continuar? (S/N): ").upper()

            if opcion == "N":

                self.partida.resultado = "Abandono"

                break

        if self.partida.resultado == "":

            self.partida.resultado = "Derrota-Sin Oxigeno"

        print("----------------------")
        print("RESULTADO:", self.partida.resultado)

        self.estadisticas.guardarPartida(
            self.jugador,
            self.partida
        )

        self.estadisticas.resultadosPartida()

        self.partida = ""

    def ejecutarJuego(self):

        while True:

            try:

                print("====================")
                print("MISION SUBMARINA")
                print("====================")
                print("1) Instrucciones del juego")
                print("2) Registrar jugador")
                print("3) Nuevo juego")
                print("4) Iniciar juego")
                print("5) Ver estadisticas")
                print("6) Ver graficos")
                print("7) Ver graficos avanzados")
                print("8) Salir")

                opcion = int(input("Ingrese opcion: "))

                if opcion == 1:

                    self.mostrarReglas()

                elif opcion == 2:

                    self.registrarJugador()

                elif opcion == 3:

                    self.nuevoJuego()

                elif opcion == 4:

                    self.iniciarJuego()

                elif opcion == 5:

                    self.estadisticas.mostrarDatos()

                elif opcion == 6:

                    self.graficos.ReporteGraficos()

                elif opcion == 7:

                    self.graficos.ReporteGraficosAvanzados()

                elif opcion == 8:

                    print("--------------------------------")
                    print("--------GRACIAS POR JUGAR-------")
                    print("--------------------------------")

                    break

                else:

                    print("--------------------------------")
                    print("-------OPCION INVALIDA----------")
                    print("--------------------------------")

            except ValueError:

                print("Dato no valido, pruebe de nuevo")



juego = Juego()
juego.ejecutarJuego()