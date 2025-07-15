import tkinter as tk
from modeloa import iniciar_app as iniciar_a
from modelob import iniciar_app as iniciar_b
from modeloc import iniciar_app as iniciar_c

def main_menu():
    root = tk.Tk()
    root.title("Menu Principal - Selecione o Modelo")

    tk.Label(root, text="Escolha o modelo de planilha para processar:", font=("Arial", 12)).pack(pady=15)

    tk.Button(root, text="Modelo A",bg='blue',fg="white",width=30, command=lambda: [root.destroy(), iniciar_a()]).pack(pady=5)
    tk.Button(root, text="Modelo B",bg='red',fg="white", width=30, command=lambda: [root.destroy(), iniciar_b()]).pack(pady=5)
    tk.Button(root, text="Modelo C",bg='White',fg="black", width=30, command=lambda: [root.destroy(), iniciar_c()]).pack(pady=5)
    root.mainloop()

if __name__ == "__main__":
    main_menu()
