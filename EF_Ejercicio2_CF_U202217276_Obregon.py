from abc import ABC, abstractmethod
import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

class PersonalSalud(ABC):
    def __init__(self, id_empleado, nombre, sueldo_base):
        self.__id_empleado = id_empleado
        self.__nombre = nombre
        self.__sueldo_base = sueldo_base

    @property
    def id_empleado(self):
        return self.__id_empleado

    @id_empleado.setter
    def id_empleado(self, valor):
        self.__id_empleado = valor

    @property
    def nombre(self):
        return self.__nombre

    @nombre.setter
    def nombre(self, valor):
        self.__nombre = valor

    @property
    def sueldo_base(self):
        return self.__sueldo_base

    @sueldo_base.setter
    def sueldo_base(self, valor):
        if valor >= 0:
            self.__sueldo_base = valor

    @abstractmethod
    def calcular_sueldo_neto(self):
        pass


class MedicoEspecialista(PersonalSalud):
    def __init__(self, id_empleado, nombre, sueldo_base, horas_cirugia, tarifa_hora):
        super().__init__(id_empleado, nombre, sueldo_base)
        self.__horas_cirugia = horas_cirugia
        self.__tarifa_hora = tarifa_hora

    @property
    def horas_cirugia(self): 
        return self.__horas_cirugia

    @horas_cirugia.setter
    def horas_cirugia(self, valor): 
        self.__horas_cirugia = valor

    @property
    def tarifa_hora(self): 
        return self.__tarifa_hora

    @tarifa_hora.setter
    def tarifa_hora(self, valor): 
        self.__tarifa_hora = valor

    def calcular_sueldo_neto(self):
        return self.sueldo_base + (self.__horas_cirugia * self.__tarifa_hora)


class EnfermeroGeneral(PersonalSalud):
    def __init__(self, id_empleado, nombre, sueldo_base, guardias_nocturnas):
        super().__init__(id_empleado, nombre, sueldo_base)
        self.__guardias_nocturnas = guardias_nocturnas

    @property
    def guardias_nocturnas(self): 
        return self.__guardias_nocturnas

    @guardias_nocturnas.setter
    def guardias_nocturnas(self, valor): 
        self.__guardias_nocturnas = valor

    def calcular_sueldo_neto(self):
        return self.sueldo_base + (self.__guardias_nocturnas * 85.0)

class GestorNominas():
    def __init__(self):
        self.__personal = []
        self.__archivo_excel = "planilla_clinicasalud.xlsx"
        self.__cargar_datos()

    def __cargar_datos(self):
        if os.path.exists(self.__archivo_excel):
            try:
                df = pd.read_excel(self.__archivo_excel)
                for _, row in df.iterrows():
                    tipo = row['Tipo']
                    id_emp = row['ID Empleado']
                    nombre = row['Nombre']
                    sueldo = float(row['Sueldo Base'])
                    
                    if tipo == 'MedicoEspecialista':
                        horas = int(row['Horas/Guardias'])
                        tarifa = float(row['Tarifa/Extra'])
                        self.__personal.append(MedicoEspecialista(id_emp, nombre, sueldo, horas, tarifa))
                    elif tipo == 'EnfermeroGeneral':
                        guardias = int(row['Horas/Guardias'])
                        self.__personal.append(EnfermeroGeneral(id_emp, nombre, sueldo, guardias))
            except Exception as e:
                print(f"Error al cargar nóminas: {e}")

    def registrar(self, persona: PersonalSalud):
        if any(p.id_empleado == persona.id_empleado for p in self.__personal):
            print("Error: Ya existe un empleado con este ID.")
            return False
        self.__personal.append(persona)
        self.exportar_excel()
        return True

    def mostrar(self):
        if not self.__personal:
            print("La planilla mensual está vacía.")
            return

        print(f"{'TIPO':<18} | {'ID EMP':<10} | {'NOMBRE':<20} | {'SUELDO BASE':<15} | {'SUELDO NETO':<15}")
        print("-" * 85)
        for p in self.__personal:
            tipo = p.__class__.__name__
            neto = p.calcular_sueldo_neto()
            print(f"{tipo:<18} | {p.id_empleado:<10} | {p.nombre:<20} | S/. {p.sueldo_base:>11.2f} | S/. {neto:>11.2f}")

    def actualizar(self, id_emp: str, nuevo_sueldo_base: float) -> bool:
        for p in self.__personal:
            if p.id_empleado == id_emp:
                p.sueldo_base = nuevo_sueldo_base
                self.exportar_excel()
                return True
        return False

    def eliminar(self, id_emp: str) -> bool:
        for p in self.__personal:
            if p.id_empleado == id_emp:
                self.__personal.remove(p)
                self.exportar_excel()
                return True
        return False

    def exportar_excel(self):
        datos = []
        for p in self.__personal:
            tipo = p.__class__.__name__
            horas_guardias = p.horas_cirugia if tipo == 'MedicoEspecialista' else p.guardias_nocturnas
            tarifa_extra = p.tarifa_hora if tipo == 'MedicoEspecialista' else 0.0
            
            datos.append({
                "Tipo": tipo,
                "ID Empleado": p.id_empleado,
                "Nombre": p.nombre,
                "Sueldo Base": p.sueldo_base,
                "Horas/Guardias": horas_guardias,
                "Tarifa/Extra": tarifa_extra,
                "Sueldo Neto": p.calcular_sueldo_neto()
            })
        df = pd.DataFrame(datos)
        df.to_excel(self.__archivo_excel, index=False)

    def generar_graficos(self):
        if not self.__personal:
            print("No hay personal registrado para graficar la distribución presupuestaria.")
            return

        categorias = np.array([p.__class__.__name__ for p in self.__personal])
        sueldos_netos = np.array([p.calcular_sueldo_neto() for p in self.__personal])

        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
        fig.suptitle("Uso Presupuestal y Control de Planilla - ClínicaSalud", fontsize=16, fontweight="bold")

        df_temp = pd.DataFrame({'Categoria': categorias, 'Sueldo': sueldos_netos})
        totales_categoria = df_temp.groupby('Categoria')['Sueldo'].sum()

        ax1.bar(totales_categoria.index, totales_categoria.values, color=['darkcyan', 'mediumpurple'], edgecolor='black', width=0.5)
        ax1.set_ylabel('Total Egresado (S/.)')
        ax1.set_title('Presupuesto Total Egresado por Categoría')
        ax1.grid(axis='y', linestyle=':', alpha=0.6)

        total_nomina = np.sum(sueldos_netos)
        ax2.pie(totales_categoria.values, labels=totales_categoria.index, autopct='%1.1f%%', startangle=90,
                colors=['cadetblue', 'orchid'], wedgeprops={"edgecolor": "black"})
        ax2.set_title(f'Distribución Porcentual Presupuestaria\n(Nómina Total: S/. {total_nomina:,.2f})')

        plt.tight_layout()
        plt.show()
    def ejecutar_menu(self):
        while True:
            print("\n========================================================")
            print("         PLANILLA DE NOMINA - CLINICASALUD              ")
            print("========================================================")
            print("1. Registrar Personal Médico/Asistencial")
            print("2. Mostrar Planilla Consolidada")
            print("3. Actualizar Sueldo Base")
            print("4. Eliminar Empleado")
            print("5. Control de Uso Presupuestal (Gráficas)")
            print("6. Salir")
            print("========================================================")
            
            try:
                opcion = int(input("Seleccione una opción: "))
                if opcion == 1:
                    print("\n----------------- NUEVO INGRESO -----------------")
                    tipo = input("Cargo (1: Médico Especialista, 2: Enfermero General): ")
                    id_emp = input("ID Empleado (Ej: M001, E002): ").upper().strip()
                    nombre = input("Nombre completo: ").strip()
                    sueldo = float(input("Sueldo Base (S/.): "))
                    
                    if tipo == '1':
                        horas = int(input("Horas de cirugía mensuales: "))
                        tarifa = float(input("Tarifa por hora de cirugía (S/.): "))
                        p = MedicoEspecialista(id_emp, nombre, sueldo, horas, tarifa)
                    elif tipo == '2':
                        guardias = int(input("Número de guardias nocturnas realizadas: "))
                        p = EnfermeroGeneral(id_emp, nombre, sueldo, guardias)
                    else:
                        print("Opción inválida.")
                        continue

                    if self.registrar(p):
                        print("Colaborador registrado de manera óptima.")
                    
                elif opcion == 2:
                    print("\n----------------- REPORTE CONTABLE DE PLANILLA -----------------")
                    self.mostrar()
                    
                elif opcion == 3:
                    print("\n----------------- ACTUALIZACIÓN REMUNERATIVA -----------------")
                    id_emp = input("Ingrese ID del Empleado a actualizar: ").upper().strip()
                    nuevo_s = float(input("Nuevo Sueldo Base (S/.): "))
                    if self.actualizar(id_emp, nuevo_s):
                        print("Sueldo base actualizado e ingresado a base de datos.")
                    else:
                        print("Colaborador no encontrado.")
                        
                elif opcion == 4:
                    print("\n----------------- ELIMINACIÓN/BAJA DE PERSONAL -----------------")
                    id_emp = input("Ingrese ID del Empleado a retirar: ").upper().strip()
                    if self.eliminar(id_emp):
                        print("Registro del colaborador removido de la planilla contable.")
                    else:
                        print("Colaborador no encontrado.")
                        
                elif opcion == 5:
                    print("\n---------------- GENERANDO BALANCE PRESUPUESTARIO GENERAL ----------------")
                    self.generar_graficos()
                    
                elif opcion == 6:
                    print("\nCerrando sistema... ¡Buen servicio!")
                    break
            except ValueError:
                print("Formato de dato inválido.")
gestor = GestorNominas()
gestor.ejecutar_menu()