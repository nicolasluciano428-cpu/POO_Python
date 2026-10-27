from abc import ABC, abstractmethod
import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

class ActivoFinanciero(ABC):
    def __init__(self, ticker, monto_invertido):
        self.__ticker = ticker
        self.__monto_invertido = monto_invertido

    @property
    def ticker(self):
        return self.__ticker

    @ticker.setter
    def ticker(self, valor):
        self.__ticker = valor

    @property
    def monto_invertido(self):
        return self.__monto_invertido

    @monto_invertido.setter
    def monto_invertido(self, valor):
        if valor >= 0:
            self.__monto_invertido = valor

    @abstractmethod
    def calcular_rendimiento(self):
        pass


class AccionMercado(ActivoFinanciero):
    def __init__(self, ticker, monto_invertido, dividendo_anual):
        super().__init__(ticker, monto_invertido)
        self.__dividendo_anual = dividendo_anual

    @property
    def dividendo_anual(self):
        return self.__dividendo_anual

    @dividendo_anual.setter
    def dividendo_anual(self, valor):
        if valor >= 0:
            self.__dividendo_anual = valor

    def calcular_rendimiento(self):
        return (self.monto_invertido * 0.07) + self.__dividendo_anual


class CriptoMoneda(ActivoFinanciero):
    def __init__(self, ticker, monto_invertido, beta_volatilidad):
        super().__init__(ticker, monto_invertido)
        self.__beta_volatilidad = beta_volatilidad

    @property
    def beta_volatilidad(self):
        return self.__beta_volatilidad

    @beta_volatilidad.setter
    def beta_volatilidad(self, valor):
        if valor > 0:
            self.__beta_volatilidad = valor

    def calcular_rendimiento(self):
        return self.monto_invertido * (0.15 * self.__beta_volatilidad)

class GestorPortafolio:
    def __init__(self):
        self.__activos = []
        self.__archivo_excel = "portafolio_finanzaspro.xlsx"
        self.__cargar_datos()

    def __cargar_datos(self):
        if os.path.exists(self.__archivo_excel):
            try:
                df = pd.read_excel(self.__archivo_excel)
                for _, row in df.iterrows():
                    tipo = row['Tipo']
                    ticker = row['Ticker']
                    monto = float(row['Monto Invertido'])
                    
                    if tipo == 'AccionMercado':
                        dividendo = float(row['Especificacion'])
                        self.__activos.append(AccionMercado(ticker, monto, dividendo))
                    elif tipo == 'CriptoMoneda':
                        beta = float(row['Especificacion'])
                        self.__activos.append(CriptoMoneda(ticker, monto, beta))
            except Exception as e:
                print(f"Error al cargar portafolio: {e}")

    def registrar(self, activo: ActivoFinanciero):
        if any(a.ticker == activo.ticker for a in self.__activos):
            print("Error: Ya existe este activo financiero")
            return False
        self.__activos.append(activo)
        self.exportar_excel()
        return True

    def mostrar(self):
        if not self.__activos:
            print("El portafolio de inversión está vacío.")
            return

        print(f"{'TIPO':<18} | {'TICKER':<10} | {'MONTO INVERTIDO':<18} | {'RENDIMIENTO NETO':<18}")
        print("-" * 75)
        for a in self.__activos:
            tipo = a.__class__.__name__
            rendimiento = a.calcular_rendimiento()
            print(f"{tipo:<18} | {a.ticker:<10} | $ {a.monto_invertido:>15.2f} | $ {rendimiento:>15.2f}")

    def actualizar(self, ticker, nuevo_monto):
        for a in self.__activos:
            if a.ticker == ticker:
                a.monto_invertido = nuevo_monto
                self.exportar_excel()
                return True
        return False

    def eliminar(self, ticker):
        for a in self.__activos:
            if a.ticker == ticker:
                self.__activos.remove(a)
                self.exportar_excel()
                return True
        return False

    def exportar_excel(self):
        datos = []
        for a in self.__activos:
            tipo = a.__class__.__name__
            especificacion = a.dividendo_anual if tipo == 'AccionMercado' else a.beta_volatilidad
            datos.append({
                "Tipo": tipo,
                "Ticker": a.ticker,
                "Monto Invertido": a.monto_invertido,
                "Especificacion": especificacion,
                "Rendimiento": a.calcular_rendimiento()
            })
        df = pd.DataFrame(datos)
        df.to_excel(self.__archivo_excel, index=False)

    def generar_graficos(self):
        if not self.__activos:
            print("No hay activos suficientes para generar proyecciones.")
            return

        tickers = np.array([a.ticker for a in self.__activos])
        montos = np.array([a.monto_invertido for a in self.__activos])
        rendimientos = np.array([a.calcular_rendimiento() for a in self.__activos])

        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
        fig.suptitle("Análisis Financiero de Portafolio - FinanzasPro", fontsize=16, fontweight="bold")

        x = np.arange(len(tickers))
        width = 0.35
        ax1.bar(x - width/2, montos, width, label='Monto Invertido', color='teal', edgecolor='black')
        ax1.bar(x + width/2, rendimientos, width, label='Rendimiento Neto Proyectado', color='darkorange', edgecolor='black')
        ax1.set_ylabel('Valor Comercial ($)')
        ax1.set_title('Monto Invertido vs. Rendimiento Estimado')
        ax1.set_xticks(x)
        ax1.set_xticklabels(tickers)
        ax1.legend()
        ax1.grid(axis='y', linestyle=':', alpha=0.6)

        años = np.array([1, 2, 3, 4, 5])
        total_invertido = np.sum(montos)
        total_rendimiento = np.sum(rendimientos)
        proyeccion_crecimiento = total_invertido + (total_rendimiento * años)

        ax2.plot(años, proyeccion_crecimiento, marker='o', linestyle='-', color='forestgreen', linewidth=2.5, label='Valor del Portafolio')
        ax2.set_xlabel('Horizonte de Tiempo (Años)')
        ax2.set_ylabel('Capital Total Proyectado ($)')
        ax2.set_title('Proyección de Rentabilidad del Portafolio')
        ax2.set_xticks(años)
        ax2.grid(True, linestyle=':')
        ax2.legend()

        plt.tight_layout()
        plt.show()

    def ejecutar_menu(self):
        while True:
            print("\n========================================================")
            print("         PORTAFOLIO FINTECH - FINANZASPRO              ")
            print("========================================================")
            print("1. Registrar Activo Financiero")
            print("2. Mostrar Balance del Portafolio")
            print("3. Actualizar Monto Invertido")
            print("4. Eliminar Activo")
            print("5. Reporte de Rendimiento Temporal (Gráficas)")
            print("6. Salir")
            print("========================================================")
            
            try:
                opcion = int(input("Seleccione una opción: "))
                if opcion == 1:
                    print("\n----------------- COMPRA DE ACTIVO -----------------")
                    tipo = input("Categoría (1: Acción tradicional, 2: Criptomoneda): ")
                    ticker = input("Ticker/Símbolo (Ej: AAPL, BTC): ").upper().strip()
                    monto = float(input("Monto Invertido ($): "))
                    
                    if tipo == '1':
                        dividendo = float(input("Dividendo Anual Estimado ($): "))
                        a = AccionMercado(ticker, monto, dividendo)
                    elif tipo == '2':
                        beta = float(input("Beta de Volatilidad (Riesgo): "))
                        a = CriptoMoneda(ticker, monto, beta)
                    else:
                        print("Opción inválida.")
                        continue

                    if self.registrar(a):
                        print("Activo integrado a la cartera de inversiones.")
                    
                elif opcion == 2:
                    print("\n----------------- VALORIZACIÓN DE CARTERA -----------------")
                    self.mostrar()
                    
                elif opcion == 3:
                    print("\n----------------- MODIFICAR INVERSIÓN -----------------")
                    ticker = input("Ingrese Ticker del activo a actualizar: ").upper().strip()
                    nuevo_m = float(input("Nuevo Monto Invertido ($): "))
                    if self.actualizar(ticker, nuevo_m):
                        print("Capital actualizado y recalculado.")
                    else:
                        print("Activo no encontrado.")
                        
                elif opcion == 4:
                    print("\n----------------- VENDER/LIQUIDAR ACTIVO -----------------")
                    ticker = input("Ingrese Ticker del activo a liquidar: ").upper().strip()
                    if self.eliminar(ticker):
                        print("Activo liquidado y retirado de la cartera activa.")
                    else:
                        print("Activo no encontrado.")
                        
                elif opcion == 5:
                    print("\n---------------- CONSOLIDANDO PROYECCIONES FINANCIERAS ----------------")
                    self.generar_graficos()
                    
                elif opcion == 6:
                    print("\nCerrando terminal...")
                    break
            except ValueError:
                print("Formato de dato inválido.")

menu=GestorPortafolio()
menu.ejecutar_menu()