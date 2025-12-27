# Induction Motor Dynamics Simulator

Web-based simulator for dynamic modeling and acceleration time estimation of double-cage induction motors. This project provides an interactive platform to analyze torque, current, slip and power behavior using equivalent circuit models and numerical methods.

---

## 📌 Project Overview

This application was developed to support engineers, researchers and students in the analysis of induction motors under different operating and starting conditions. The simulator implements mathematical models based on the electrical equivalent circuit and slip-dependent dynamics to reproduce the real behavior of three-phase induction motors.

The system allows users to:
- Register and manage motor parameters  
- Select starting methods (direct, star-delta, autotransformer)  
- Simulate motor torque, current and power curves  
- Model different load profiles (constant, linear, quadratic and cubic)  
- Estimate the total acceleration time from standstill to nominal speed  

All results are displayed through interactive charts and tables, enabling both visual and numerical analysis.

---

## 🧠 Mathematical Model

The motor behavior is calculated from:
- Stator and rotor equivalent impedances  
- Thévenin equivalent voltage and impedance  
- Slip-dependent current and torque equations  
- Dynamic resistance variation for double-cage rotor modeling  
- Torque balance between motor and load  
- Acceleration time obtained from the integration of torque and inertia  

This approach allows accurate prediction of motor starting performance in industrial scenarios.

---

## 🛠 Technologies Used

- **Python 3**
- **Flask** – Web framework  
- **NumPy** – Numerical computation  
- **SQLAlchemy** – Database ORM  
- **SQLite / PostgreSQL** – Data storage  
- **Plotly** – Interactive plotting  
- **HTML, CSS, JavaScript** – Frontend interface  

---

## ⚙️ Installation

```bash
git clone https://github.com/yourusername/induction-motor-dynamics-simulator.git
cd induction-motor-dynamics-simulator
pip install -r requirements.txt
python app.py
