# coding: utf-8

import win32com.client
import os
from pylab import *
from sys import exit

class DSS:

    def __init__(self, file_name):

        self.path_dss = self.get_dss_files_dir()  + file_name      

        # Create connection Python - OpenDSS
        self.dssObj = win32com.client.Dispatch("OpenDSSEngine.DSS")

        # Start OpenDSS Object
        if self.dssObj.Start(0) == False:
            print ("Problems starting OpenDSS...")
            exit()
        else:
            # Create variables for the main interfaces
            self.dssText = self.dssObj.Text
            self.dssCircuit = self.dssObj.ActiveCircuit
            self.dssSolution = self.dssCircuit.Solution
            self.dssCktElement = self.dssCircuit.ActiveCktElement
            self.dssBus = self.dssCircuit.ActiveBus
            self.dssLines = self.dssCircuit.Lines
            self.dssTransformers = self.dssCircuit.Transformers
    
    
    @staticmethod        
    def get_dss_files_dir() -> str:
      cwd = os.getcwd()
      return cwd + r"\\DSS\\"
      
      
    def versao_DSS(self) -> str:

        return self.dssObj.Version

    def compile_DSS(self) -> None:

        # Clear last simulation data
        self.dssObj.ClearAll()

        self.dssText.Command = "compile " + self.path_dss

    def solve_DSS_snapshot(self, load_mult) -> None:

        # Settings
        self.dssText.Command = "Set Mode=SnapShot"
        self.dssText.Command = "Set ControlMode=Static"

        # Multiply the nominal value of the loads by the load_mult
        self.dssSolution.LoadMult = load_mult

        # Solves the Power Flow
        self.dssSolution.Solve()

    def get_power_results(self) -> None:
        self.dssText.Command = "Show powers kva elements"

    def get_circuit_name(self) -> str:
        return self.dssCircuit.Name

    def get_circuit_power(self) -> tuple:

        p = -self.dssCircuit.TotalPower[0]
        q = -self.dssCircuit.TotalPower[1]

        return p, q

    def active_bus(self, bus_name:str) -> str:

        # Activate the bus by your name
        self.dssCircuit.SetActiveBus(bus_name)

        # Return activated bus by its name
        return self.dssBus.Name

    def get_bus_distance(self) -> float:
        return self.dssBus.Distance

    def get_bus_kVBase(self) -> float:
        return self.dssBus.kVBase

    def get_bus_VMagAng(self) -> float:
        return self.dssBus.VMagAngle

    def activate_element(self, element_name: str) -> str:

        # Activates element by its full name Type.Name
        self.dssCircuit.SetActiveElement(element_name)

        # Returns activated element name
        return self.dssCktElement.Name

    def get_element_bus(self) -> tuple:

        bus = self.dssCktElement.BusNames

        bus1 = bus[0]
        bus2 = bus[1]

        return bus1, bus2

    def get_element_voltage(self) -> float:

        return self.dssCktElement.VoltagesMagAng

    def get_element_power(self) -> float:
        return self.dssCktElement.Powers

    def get_line_name(self) -> str:
        return self.dssLines.Name

    def get_line_length(self) -> float:
        return self.dssLines.Length

    def set_line_length(self, length: float):
        self.dssLines.Length = length

    def get_transformer_name(self) -> str:
        return self.dssTransformers.Name

    def get_terminal_voltage_transformer(self, terminal: int) -> float:

        # Activare transformer terminals
        self.dssTransformers.Wdg = terminal

        return self.dssTransformers.kV

    def get_line_name_and_length(self):

        lines_name_list = []
        lines_length_list = []

        # Seleciona a primeira linha
        self.dssLines.First

        for i in range(self.dssLines.Count):

            lines_name_list.append(self.dssLines.Name)
            lines_length_list.append(self.dssLines.Length)

            self.dssLines.Next

        return lines_name_list, lines_length_list


if __name__ == "__main__":
  
    # Criar um objeto da classe DSS
    objeto = DSS("index.dss")

    print ("Versão do OpenDSS: " + objeto.versao_DSS() + "\n")

    # Resolver o Fluxo de Potência
    objeto.compile_DSS()
    objeto.solve_DSS_snapshot(1.0)

    # Arquivo de Resultado
    objeto.get_power_results()

    # Informações do elemento Circuit
    p, q = objeto.get_circuit_power()
    print ("Nosso exemplo apresenta o nome do elemnto Circuit: " + objeto.get_circuit_name())
    print ("Fornece Potência Ativa: " + str(p) + " kW")
    print ("Fornece Potência Reativa: " + str(q) + " kvar \n")

    # Informações da Barra escolhida
    print ("Barra Ativa: " + objeto.active_bus("C"))
    print ("Distância do EnergyMeter: " + str(objeto.get_bus_distance()))
    print ("Tensão de Base da Barra (kV) : " + str(sqrt(3) * objeto.get_bus_kVBase()))
    print ("Tensões dessa Barra (kV): " + str(objeto.get_bus_VMagAng()) + "\n")

    # Informações do elemento escolhido
    print ("Elemento Ativo: " + objeto.activate_element("Line.Linha1"))
    barra1, barra2 = objeto.get_element_bus()
    print ("Esse elemento está conectado entre as barras: " + barra1 + " e " + barra2)
    print ("As tensões nodais desse elemento (kV): " + str(objeto.get_element_voltage()))
    print ("As potências desse elemento (kW) e (kvar): " + str(objeto.get_element_power()) + "\n")

    # Informações dos dados da linha escolhida
    print ("Elemento Ativo: " + objeto.activate_element("Line.Linha1"))
    print ("Nome da linha ativa: " + objeto.get_line_name())
    print ("Tamanho da linha ativa: " + str(objeto.get_line_length()))
    print ("Alterando o tamanho da linha para 0.4 km.")
    objeto.set_line_length(0.4)
    print ("Novo tamanho da linha ativa: " + str(objeto.get_line_length()) + "\n")

    # informações do transformador
    print ("Transformador Ativo: " + objeto.activate_element("Transformer.Trafo"))
    print ("Nome do Transformador ativo: " + objeto.get_transformer_name())
    print ("Tensão nominal do primário: " + str(objeto.get_terminal_voltage_transformer(1)))
    print ("Tensão nominal do secundário: " + str(objeto.get_terminal_voltage_transformer(2)) + "\n")

    # Nome e tamanho de todas as linhas
    nome_linhas, tamanho_linhas = objeto.get_line_name_and_length()
    print ("Nomes das linhas: " + str(nome_linhas))
    print ("Tamanhos das linhas: " + str(tamanho_linhas))