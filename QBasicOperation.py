#!/usr/bin/env python
# coding: utf-8

Debug = 0

import time

if Debug:
    print('Begin')
    time_before = time.time()
    
import functools
from re import sub as re_sub
from numpy import arange as np_arange
from sympy.parsing import sympy_parser
from decimal import Decimal, DecimalTuple
from fractions import Fraction
from _functools import reduce
from traceback import format_exc

if Debug:
    print('Finish importation-ordinary', round(time.time()-time_before, 5))
    time_before = time.time()

try: from qiskit import qpy
except ImportError: pass

from qiskit import QuantumCircuit
if Debug:
    print('Finish importation-Qiskit-0', round(time.time()-time_before, 5))
    time_before = time.time()
from qiskit_aer import Aer, AerSimulator

if Debug:
    print('Finish importation-Qiskit-1', round(time.time()-time_before, 5))
    time_before = time.time()

from qiskit.circuit.library import WeightedAdder, RGQFTMultiplier, VBERippleCarryAdder

if Debug:
    print('Finish importation-Qiskit-2', round(time.time()-time_before, 5))
    time_before = time.time()

from qiskit.circuit import Instruction, CircuitInstruction, Qubit, QuantumRegister, Clbit, ClassicalRegister
from qiskit.circuit.library.standard_gates import IGate, XGate, CXGate, CCXGate, C3XGate, C4XGate, MCXGate, RXGate, RYGate, RZGate, HGate
from qiskit.exceptions import QiskitError

if Debug:
    print('Finish importation-Qiskit-3', round(time.time()-time_before, 5))
    time_before = time.time()

from qiskit import QuantumRegister
from qiskit import ClassicalRegister

if Debug:
    print('Finish importation-Qiskit-4', round(time.time()-time_before, 5))
    time_before = time.time()

from qiskit import transpile
# from qiskit import assemble
from qiskit.circuit.library import DraperQFTAdder
from qiskit.converters import circuit_to_dag, dag_to_circuit
from qiskit.transpiler import PassManager
from qiskit.transpiler.passes import Decompose
# from qiskit.utils import QuantumInstance

if Debug:
    print('Finish importation-Qiskit', round(time.time()-time_before, 5))
    time_before = time.time()

def writeCircuit(qc, filename):
    with open(filename, 'wb') as fd:
        qpy.dump(qc, fd)
        
def readCircuit(filename):
    with open(filename, 'rb') as fd:
        new_qc = qpy.load(fd)[0]
        return new_qc
        
def dec_to_bin(decimal, num_bits):
    binary = bin(decimal)[2:].zfill(num_bits)
    return [int(bit) for bit in binary]

def bin_to_dec(binary):
    binary_str = ''.join(map(str, binary))
    return int(binary_str, 2)

def toIntAdder(a, b):
    # Convert to strings to manipulate them easily
    a_str = str(a)
    b_str = str(b)

    # Find the length of the decimal part
    aList = list(Decimal(a_str).as_tuple())
    bList = list(Decimal(b_str).as_tuple())

    # Find the minimum decimal exponent/place values
    placeVal = min(aList[2], bList[2])

    # Shift the decimal point to the rightmost position
    aList[2] -= placeVal
    bList[2] -= placeVal
    
    # Store Final Values
    a_shifted = int(Decimal(aList))
    b_shifted = int(Decimal(bList))

    return a_shifted, b_shifted, placeVal

def toIntMultiplier(a, b):
    # Convert to strings to manipulate them easily
    a_str = str(a)
    b_str = str(b)

    # Find the length of the decimal part
    aList = list(Decimal(a_str).as_tuple())
    bList = list(Decimal(b_str).as_tuple())

    # Find the minimum decimal exponent/place values
    placeVal = aList[2] + bList[2]

    # Shift the decimal point to the rightmost position
    aList[2] = 0
    bList[2] = 0
    
    # Store Final Values
    a_shifted = int(Decimal(aList))
    b_shifted = int(Decimal(bList))

    return a_shifted, b_shifted, placeVal
    
def toFloat(a, placeVal):
    # Convert to strings to manipulate them easily
    a_str = str(a)

    # Find the length of the decimal part
    aList = list(Decimal(a_str).as_tuple())

    # Shift the decimal point to the leftmost position
    aList[2] += placeVal
    
    # Store Final Values
    a_shifted = float(Decimal(aList))

    return a_shifted
    
class QAnd:
    def __init__(self, debug=None):
        self.debug = debug
        
        # Create a Quantum Circuit with 3 qubits
        q = QuantumRegister(3,'q')
        c = ClassicalRegister(1,'c')
        self.circuit = circuit = QuantumCircuit(q,c)
        beginTime = time.time()
        
        # q, c: list of qubits and clbits
        self.q = circuit.qubits
        self.c = circuit.clbits
        
        # Initialize identity gates
        for i in range(2):
            circuit.id(i)

        # Apply the CCX gate (Toffoli gate) on qubit 0, 1, and use qubit 2 as the target qubit
        circuit.ccx(0, 1, 2)

        # Measure the third qubit
        circuit.measure(2, 0)
        
        endTime = time.time()
        # print(f'QAdd-Direct-Time Elapsed = {endTime-beginTime}')
        
    def exec(self, i1, i2, boolEnhance=False, debug=None):
        args = i1, i2
        circuit = self.circuit
        
        for i in range(2):
            circuit.data[i] = CircuitInstruction(operation=Instruction(name='x' if bool(int(bool(args[i]))) else 'id', num_qubits=1, num_clbits=0, params=[]), qubits=[self.q[i]], clbits=())

        # Visualize the circuit
        # print(circuit.draw())

        # Simulate the circuit
        simulator = Aer.get_backend('qasm_simulator')
        compiled_circuit = transpile(circuit, simulator)
        # qobj = assemble(compiled_circuit)
        result = simulator.run(compiled_circuit).result()

        # Get the result
        counts = result.get_counts(circuit)
        value = list(counts)[0]
        
        return bool(int(value)) if boolEnhance else int(value)
        
if Debug:
    print('Finish defs QAnd')
    
class QOr:
    def __init__(self, debug=None):
        self.debug = debug
        
        # Create a Quantum Circuit with 3 qubits
        q = QuantumRegister(3,'q')
        c = ClassicalRegister(1,'c')
        self.circuit = circuit = QuantumCircuit(q,c)
        beginTime = time.time()
        
        # q, c: list of qubits and clbits
        self.q = circuit.qubits
        self.c = circuit.clbits
        
        # Initialize identity gates
        for i in range(2):
            circuit.id(i)

        # Apply the CCX gate (Toffoli gate) on qubit 0, 1, and use qubit 2 as the target qubit
        circuit.ccx(0, 1, 2)
        
        # Apply final Reversal gate
        circuit.x(2)
        
        # Measure the third qubit
        circuit.measure(2, 0)
        
        endTime = time.time()
        # print(f'QAdd-Direct-Time Elapsed = {endTime-beginTime}')
        
    def exec(self, i1, i2, boolEnhance=False, debug=None):
        args = i1, i2
        circuit = self.circuit
        
        for i in range(2):
            circuit.data[i] = CircuitInstruction(operation=Instruction(name='id' if bool(int(bool(args[i]))) else 'x', num_qubits=1, num_clbits=0, params=[]), qubits=[self.q[i]], clbits=())

        # Visualize the circuit
        # print(circuit.draw())

        # Simulate the circuit
        simulator = Aer.get_backend('qasm_simulator')
        compiled_circuit = transpile(circuit, simulator)
        # qobj = assemble(compiled_circuit)
        result = simulator.run(compiled_circuit).result()

        # Get the result
        counts = result.get_counts(circuit)
        value = list(counts)[0]
        
        return bool(int(value)) if boolEnhance else int(value)
        
if Debug:
    print('Finish defs QOr')

class QXOr:
    def __init__(self, file=None, debug=None):
        self.file = file
        self.debug = debug
        
        beginTime = time.time()
        
        if file:
            self.circuit = circuit = readCircuit(file)
            
        else:    
            # Create a Quantum Circuit with 2 qubits
            q = QuantumRegister(2,'q')
            c = ClassicalRegister(1,'c')
            self.circuit = circuit = QuantumCircuit(q,c)
            
            # Initialize identity gates
            for i in range(2):
                circuit.id(i)

            # Apply the CX gate (CNOT/Controlled-X gate) with qubit 0 as control and qubit 1 as target
            circuit.cx(0, 1)
            
            # Measure the second qubit
            circuit.measure(1, 0)
        
        # q, c: list of qubits and clbits
        self.q = circuit.qubits
        self.c = circuit.clbits
        
        endTime = time.time()
        # print(f'QAdd-Direct-Time Elapsed = {endTime-beginTime}')
        
    def exec(self, i1, i2, boolEnhance=False, debug=None):
        args = i1, i2
        circuit = self.circuit
        
        for i in range(2):
            circuit.data[i] = CircuitInstruction(operation=Instruction(name='x' if bool(int(bool(args[i]))) else 'id', num_qubits=1, num_clbits=0, params=[]), qubits=[self.q[i]], clbits=())

        # Visualize the circuit
        # print(circuit.draw())

        # Simulate the circuit
        simulator = Aer.get_backend('qasm_simulator')
        
        compiled_circuit = transpile(circuit, simulator)
        # compiled_circuit = execute(circuit, simulator)
        # quantum_instance = QuantumInstance(backend=simulator)
        
        # qobj = assemble(compiled_circuit)
        result = simulator.run(compiled_circuit).result()
        # result = quantum_instance.execute(circuit)

        # Get the result
        counts = result.get_counts(circuit)
        value = list(counts)[0]
        
        return bool(int(value)) if boolEnhance else int(value)
        
if Debug:
    print('Finish defs QXOr')

quOr = QOr()
quXOr = QXOr()
quAnd = QAnd()

and_gate = quAnd.exec
or_gate = quOr.exec
xor_gate = quXOr.exec

def half_adder(a, b):
    sum_bit = xor_gate(a, b)
    carry_bit = and_gate(a, b)
    return sum_bit, carry_bit

def full_adder(a, b, carry_in):
    sum_bit_1, carry_bit_1 = half_adder(a, b)
    sum_bit_2, carry_bit_2 = half_adder(sum_bit_1, carry_in)
    carry_out = or_gate(carry_bit_1, carry_bit_2)
    return sum_bit_2, carry_out

def multi_bit_adder(a, b):
    n = len(a)
    result = [0] * (n + 1)
    carry = 0
    for i in range(n):
        result[i], carry = full_adder(a[i], b[i], carry)
    result[n] = carry
    return result

def adderInt(decimal_a, decimal_b):
    max_bits = max(decimal_a.bit_length(), decimal_b.bit_length()) + 1  # Add 1 for the possible carry out
    binary_a = dec_to_bin(decimal_a, max_bits)
    binary_b = dec_to_bin(decimal_b, max_bits)
    
    sum_bits = []
    carry_in = 0
    for i in range(max_bits - 1, -1, -1):  # Loop through bits from most significant to least significant
        sum_bit, carry_in = full_adder(binary_a[i], binary_b[i], carry_in)
        sum_bits.insert(0, sum_bit)  # insert the sum bit at the beginning
    return bin_to_dec(sum_bits)

def multiplier(a, b):
    n = len(a)
    result = [0] * (2 * n)
    for i in range(n):
        for j in range(n):
            if b[j]:
                result[i + j:i + j + n] = multi_bit_adder(result[i + j:i + j + n], [0] * i + [a[i]])
    return result

def adder(aRaw, bRaw):
    a, b, placeVal = toIntAdder(aRaw, bRaw)
    result1 = adderInt(a, b)
    result2 = toFloat(result1, placeVal)
    result2Int = int(result2)
    return result2Int if result2Int == result2 else result2

def sum(*args):
    return functools.reduce(adder, args)

def subtractor(a, b):
    return float(a) - float(b)
    
if __name__ == '__main__':
    # Example usage
    a = 5000.55
    b = 2.2

    # for a in range(10):
        # for b in range(10):
    result = adder(a, b)

    print(a, b, "Result:", result, a+b, result==a+b)  # Output: 11, which is 5 + 6 = 11 in decimal