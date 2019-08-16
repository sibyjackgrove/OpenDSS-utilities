#OpenDSS classes

import os
import math
import numpy as np
import win32com.client


class DSS():

    def __init__(self,DSS_model_file):
        
        #Change directory
        os.chdir(os.getcwd())
        
        self.DSS_model_file = DSS_model_file

        # Start OpenDSS engine
        self._dssObj = win32com.client.Dispatch("OpenDSSEngine.DSS")

        # Check if OpenDSS engine started
        if not self._dssObj.Start(0):
            print('OpenDSS engine failed to start!')
        else:
            print('OpenDSS engine started.')
            # Create variables for main interfaces
            self.dssText = self._dssObj.Text
            self.dssCircuit = self._dssObj.ActiveCircuit
            self.dssSolution = self.dssCircuit.Solution
            self.dssCktElement = self.dssCircuit.ActiveCktElement
            self.dssBus = self.dssCircuit.ActiveBus
            self.dssLines = self.dssCircuit.Lines
            self.dssTransformers = self.dssCircuit.Transformers
    
    @property
    def version(self):

        return self._dssObj.Version    
    
    def compile_DSS(self):

        # Clear information from last simulation
        self._dssObj.ClearAll()

        self.dssText.Command = "compile " + self.DSS_model_file
        
    def solve_DSS_snapshot(self, load_multiplier=1.0):

        # Settings
        self.dssText.Command = "Set Mode=SnapShot"
        self.dssText.Command = "Set ControlMode=Static"

        # Multiply the nominal value of the loads by the load multiplier
        self.dssSolution.LoadMult = load_multiplier

        # Solve Power Flow
        self.dssSolution.Solve()

    def create_results_power(self):
        """Create the results file for power."""
        
        self.dssText.Command = "Show powers kva elements"
    
    def create_results_voltage(self):
        """Create the results file for voltages."""
        
        self.dssText.Command = "Show voltages kv elements"
        
    def get_circuit_name(self):
        """Return name of DSS circuit."""

        return self.dssCircuit.Name

    def get_circuit_power(self):
        """Return active and reactive power of cirucit."""

        P = -self.dssCircuit.TotalPower[0]
        Q = -self.dssCircuit.TotalPower[1]

        return P, Q    
    
    def show_bus_info(self,bus_name):
        """Show information about selected bus."""
        
        print("Active bus:{}".format(self.set_active_bus(bus_name)))
        print("Distance to EnergyMeter:{}".format(self.get_bus_distance()))
        print("Bus base voltage(kV):{}".format(str(math.sqrt(3) * self.get_bus_kVBase())))
        print("Bus voltage(kV):{}".format(str(self.get_bus_VMagAng())))
        
    def set_active_bus(self, bus_name):
        """Activate bus by its name."""
        
        self.dssCircuit.SetActiveBus(bus_name)

        # Return the name of the active bus
        return self.dssBus.Name

    def get_bus_distance(self):
        """Return bus distance."""
        
        return self.dssBus.Distance

    def get_bus_kVBase(self):
        """Return the kV base of the bus."""
        
        return self.dssBus.kVBase

    def get_bus_VMagAng(self):
        """Return the bug voltage magnitude and angle."""
        
        return self.dssBus.VMagAngle      
    
    def set_active_element(self, element_name):
        """Activates element by its full name Type.Name"""

        self.dssCircuit.SetActiveElement(element_name)

        # Return enabled element name
        return self.dssCktElement.Name

    def get_element_buses(self):
        """Return buses in the element."""
        
        bus_1 = self.dssCktElement.BusNames[0]
        bus_2 = self.dssCktElement.BusNames[1]

        return bus_1, bus_2

    def get_element_voltages(self):

        return self.dssCktElement.VoltagesMagAng

    def get_element_powers(self):
        
        return self.dssCktElement.Powers
    
    def get_element_name(self):
        
        return self.dssCktElement.Name    
    
    def show_element_info(self,element_name):
        """Show information about selected element."""
        
        print ('Active element:{}'.format(self.set_active_element(element_name)))
        bus1, bus2 = self.get_element_buses()
        print('This element is connected between the buses:{} and {}'.format(bus1,bus2))
        print('Nodal voltages of this element (kV):{}'.format(str(self.get_element_voltages())))
        print('Powers of this element (kW) and (kvar):{}'.format(str(self.get_element_powers())))
    
    def change_line_length(self,line_name,length):
        """Change length of chosen line."""
        
        
        print('Changing length of line:{}'.format(self.set_active_element(line_name)))
              
        print('Name of active line:{}'.format(self.get_line_name()))
        print('Current length of active line:{} km'.format(self.get_line_length()))
              
        print ("Changing line to {} km".format(length))
        self.set_line_length(length)
        print("New length of line:{} km".format(self.get_line_length()))
    
    def get_line_name(self):
        """Return line name."""
        
        return self.dssLines.Name

    def get_line_length(self):
        """Return line length."""
        
        return self.dssLines.Length

    def set_line_length(self, length):
        """Set line name."""
        
        self.dssLines.Length = length

    def show_transformer_info(self,transformer_name):
        """Show transformer information."""
        
        print('Active element:{}'.format(self.set_active_element(transformer_name)))
        
        print('Name of active Transformer:{}'.format(self.get_transformer_name()))
        
        print('Bus connected to primary winding:{}'.format(str(self.get_transformer_bus(1))))
        print('Rated voltage at primary winding:{} kV'.format(str(self.get_transformer_voltage(1))))
        print('Rated power at primary winding:{} kVA'.format(str(self.get_transformer_S(1))))
        
        print('Bus connected to secondary winding:{}'.format(str(self.get_transformer_bus(2))))
        print('Rated voltage at secondary winding:{} kV'.format(str(self.get_transformer_voltage(2))))    
        print('Rated power at secondary winding:{} kVA'.format(str(self.get_transformer_S(2))))
    
    def get_transformer_name(self):
        """Return transformers name."""
        
        return self.get_element_name() #self.dssTransformers.Name

    def get_transformer_bus(self,terminal):
        """Return transformer buses."""
        
        self.dssCktElement.Properties('wdg').Val= terminal
        
        return self.dssCktElement.Properties('bus').Val
    
    def get_transformer_voltage(self, terminal):
        """Return voltage at specied transformer terminal."""
        
        #self.dssTransformers.Wdg = terminal #Activate specified terminal
        self.dssCktElement.Properties('wdg').Val= terminal #Activate specified terminal

        return self.dssCktElement.Properties('kV').Val
    
    def get_transformer_S(self, terminal):
        """Return power at specied transformer terminal."""        
        
        self.dssCktElement.Properties('wdg').Val= terminal #Activate specified terminal

        return self.dssCktElement.Properties('kVA').Val       
  
    def check_terminal(self,terminal):
        """Check whether terminal is correct."""
        
        assert terminal == 1 or terminal == 2, "Terminal should be 1 or 2."
    
    def create_transformer(self,transformer_name,bus_name,winding1_voltage,winding2_voltage,Srated=1000,connection='wye'):
        """Create and attach a transformer.
        
        Args:
             transformer_name (str): The transformer name.
             bus_name (str): The bus to which winding 1 of the transformer will be connected to.
             winding1_voltage (float): The HV side voltage in kV. 
             winding2_voltage (float): The LV side voltage in kV.
             Srated (float): Rated power in kVA.
             connection (str): Connection type ('wye' or 'delta').
        Returns:
             str: Bus name
        """
        
        transformer_name="Transformer." + transformer_name
        
        #self.dssText.Command = "New {}  Phases=3   Windings=2  XHL=1 Wdg=1 bus=A Kv=13.8 Kva=500 Conn=wye Wdg=2 bus=A_tfr Kv=0.12 Kva=500 Conn=wye".format(transformer_name)
        self.dssText.Command = "New {}  Phases=3   Windings=2  XHL=1".format(transformer_name)
        
        self.dssCircuit.SetActiveElement(transformer_name) # make transformer as active element        
        
        for i,voltage in enumerate([winding1_voltage,winding2_voltage]):
            
            self.dssCktElement.Properties('Wdg').Val=i+1
            
            if self.dssCktElement.Properties('Wdg').Val == '1':
                self.dssCktElement.Properties('bus').Val=bus_name
                
            elif self.dssCktElement.Properties('Wdg').Val == '2':
                self.dssCktElement.Properties('bus').Val='{}_tfr'.format(bus_name)
            
            self.dssCktElement.Properties('Kv').Val=voltage
            self.dssCktElement.Properties('Kva').Val=Srated
            self.dssCktElement.Properties('Conn').Val=connection
            self.dssCktElement.Properties('%r').Val=.1
            
        print('Transformer {} created at bus {}!'.format(self.get_transformer_name(),bus_name))
        return self.dssCktElement.Properties('bus').Val
    
    
    def show_all_lines(self):
        """Show name and length of all lines."""
      
        line_names, line_lengths = self.get_lines_name_and_length()
        print('Names of lines:{}'.format(str(line_names)))
        print('Length of lines:{}'.format(str(line_lengths)))
        
    def get_lines_name_and_length(self):

        # Define lists
        line_names = []
        line_lengths = []

        # Select first line
        self.dssLines.First

        for i in range(self.dssLines.Count):

            line_names.append(self.get_line_name())
            line_lengths.append(self.get_line_length())

            self.dssLines.Next

        return line_names, line_lengths
    
    def show_all_transformers(self):
        """Show name and length of all lines."""
      
        transformer_names = self.get_transformer_names()
        print('Names of transformers:{}'.format(str(transformer_names)))        
        
    def get_transformer_names(self):

        # Define lists
        transformer_names = []
        
        # Select first line
        self.dssTransformers.First

        for i in range(self.dssTransformers.Count):

            transformer_names.append(self.get_transformer_name())
            
            self.dssTransformers.Next

        return transformer_names