import win32com.client
import os  # Importing OS functions
import numpy as np
import matplotlib.pyplot as plt
import datetime


# -------------------------------------------------------------------------------//
# Python file for OpenDSS - MV Distribution Network Model
# by Ing. S.W. Roelofs
# Electrical Energy Systems Group, Dep. of Electrical Engineering
# Eindhoven University of Technology, The Netherlands
# v.1 04/05/2021
# -------------------------------------------------------------------------------//


# %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
# -----------------------------------------------------------------------------------
# Select MV_Grid file and simulation properties
# -----------------------------------------------------------------------------------
ringMV_Grid = 1
otherPossibleMV_Grid = 2

MV_Grid = ringMV_Grid  # Select MV_Grid
singlePhase = False  # with or without neutral representation   MV_Grid

timesteps = 35040  # Set no. of timesteps for simulations (35040 = 1 full year simulation) Not needed for snapshot


class DSS():
    def __init__(self):

        if MV_Grid is ringMV_Grid:
            # Define the Path to the Model
            if not singlePhase:
                self.path_model_dss = '"' + os.getcwd() + r'\Grids\MV_Ring\threePhase\Main.dss'.format(
                    MV_Grid)  # path_model_dss
            if singlePhase:
                self.path_model_dss = '"' + os.getcwd() + r'\Grids\MV_Ring\singlePhase\Main.dss'.format(
                    MV_Grid)  # path_model_dss

        if MV_Grid is otherPossibleMV_Grid:

            if not singlePhase:
                self.path_model_dss = '"' + os.getcwd() + r'\Grids\MV_Ring\Main.dss'.format(MV_Grid)  # path_model_dss
            if singlePhase:
                self.path_model_dss = '"' + os.getcwd() + r'\Grids\MV_Ring\Main.dss'.format(MV_Grid)  # path_model_dss

        # -----------------------------------------------------------------------------------
        # 1.1 -- Create the link between Python & Open DSS
        # -----------------------------------------------------------------------------------

        self.dssObj = win32com.client.Dispatch("OpenDSSEngine.DSS")

        # make DSS object
        if not self.dssObj.Start:
            print("Trouble initiating OpenDSS")
        else:
            # Create interface for Variables of Interest
            self.dssText = self.dssObj.Text  # Writes directly in DSS software
            self.dssCircuit = self.dssObj.ActiveCircuit  # Access to Active Circuit
            self.dssBus = self.dssCircuit.ActiveBus
            self.dssCktElement = self.dssCircuit.ActiveCktElement
            self.dssLines = self.dssCircuit.Lines
            self.dssSolution = self.dssCircuit.Solution
            self.dssTransformers = self.dssCircuit.Transformers

    # %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


if __name__ == "__main__":

    print("Author: S.W. Roelofs \n")
    print("Getting results from OpenDSS using Python \n")
    MV_Grid = DSS()  # "MV_Grid" is just the variable name used though the simulation. It could be anything
    print("OpenDSS " + MV_Grid.dssObj.Version + "\n")  # Prints version of OpenDSS

    # %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

    begin_time = datetime.datetime.now()
    print('Begin time simulation: ', begin_time)  # Display begin time of simulation
    # -----------------------------------------------------------------------------------
    # BEGIN Compile and Run power flow using OpenDSS through Python.
    # -----------------------------------------------------------------------------------

    MV_Grid.dssObj.ClearAll()
    MV_Grid.dssText.Command = "compile " + MV_Grid.path_model_dss

    print("compile " + MV_Grid.path_model_dss)
    MV_Grid.dssText.Command = "Set Mode=snapshot"
    MV_Grid.dssSolution.Solve()

    if MV_Grid.dssSolution.Converged is True:
        print("Power flow Converges \n")
    else:
        print("Power flow does not converge \n")
        exit()

    MV_Grid.dssText.Command = "Show powers kva elements"
    Simulation_time = datetime.datetime.now() - begin_time
    print('Power flow runtime =', Simulation_time, "\n")
    # dictionary:
    VA = {}  # pu Voltage in phase A
    VB = {}  # pu Voltage in phase B
    VC = {}  # pu Voltage in phase C
    Pij = {}  # Total absolute active power on a line
    Qij = {}  # Total absolute reactive power on a line
    Sij = {}  # Total aparent power flow on a line
    #   Get voltages from solution
    V0 = []
    V1 = []
    V2 = []
    voltageSeq_list_bus = []
    for i in range(MV_Grid.dssCircuit.NumBuses):        #get node voltages calculation
        MV_Grid.dssCircuit.SetActiveBus(MV_Grid.dssCircuit.AllBusNames[i])
        voltageSeq_list_bus.append(MV_Grid.dssBus.SeqVoltages)
        Vs = voltageSeq_list_bus[i]
        V0.append(Vs[0] / (MV_Grid.dssBus.kVBase * 1000))
        V1.append(Vs[1] / (MV_Grid.dssBus.kVBase * 1000))
        V2.append(Vs[2] / (MV_Grid.dssBus.kVBase * 1000))
    VA = MV_Grid.dssCircuit.AllNodeVmagPUByPhase(1)
    VB = MV_Grid.dssCircuit.AllNodeVmagPUByPhase(2)
    VC = MV_Grid.dssCircuit.AllNodeVmagPUByPhase(3)
    nome_list_bus = MV_Grid.dssCircuit.AllBusNames
    print("Node voltages in phase 1", VA, "\nNode voltages in phase 2", VB, "\nNode voltages in phase 3", VC, "\n")

    #   Get total Power Losses
    total_power_loss = MV_Grid.dssCircuit.Losses

    #   Get line power flows
    for h in range (MV_Grid.dssLines.Count) :
        MV_Grid.dssCircuit.SetActiveElement('Line.LINE'+format(h + 1))
        Pij["P_line{0}".format(h + 1)] = [MV_Grid.dssCktElement.Powers[0]]
        Qij["Q_line{0}".format(h + 1)] = [MV_Grid.dssCktElement.Powers[1]]
        Sij['S_line{0}'.format(h + 1)] = [np.sqrt((Pij["P_line{0}".format(h + 1)][0]) ** 2 + (Qij["Q_line{0}".format(h + 1)][0]) ** 2)]

    p = -MV_Grid.dssCircuit.TotalPower[0]
    q = -MV_Grid.dssCircuit.TotalPower[1]
    s = np.sqrt(p ** 2 + q ** 2)
    pf = p/s


    #print total line flows
    print("Total P in each Line: \n", Pij)
    print("Total Q in each Line: \n", Qij)
    print("Total S in each Line: \n", Sij, "\n")

    #print total powers from substation and powerfactor
    print("Total P from substation: " + str(p), "kW \n")
    print("Total Q from substation: " + str(q), "kvar \n")
    print("Total S from substation: " + str(s), "kVA \n")
    print("Power factor = " + str(pf), "\n")

    #print losses
    print("Total Power losses [kW]:", total_power_loss[0]/1000)
    print("Total Reactive power losses [kVAr]:", total_power_loss[1]/1000, "\n")

    # plotting
    plt.figure(1)
    plt.subplot(1,2,1)
    plt.plot(nome_list_bus, VA, label='Phase A')
    plt.plot(nome_list_bus, VB, label='Phase B')
    plt.plot(nome_list_bus, VC, label='Phase C')
    plt.title('Voltage magnitudes')
    plt.xlabel('Bus number')
    plt.ylabel("Voltage magnitude [P.U.]")
    plt.legend()
   
    plt.subplot(1,2,2)
    plt.plot(range(len(Pij)), list(Pij.values()), label='Active power')
    plt.plot(range(len(Qij)), list(Qij.values()), label='Reactive Power')
    plt.title('Branch Power Flows')
    plt.xlabel('Line')
    plt.ylabel("Power [kW,kvar]")
    plt.legend()
    plt.show()
