import tkinter as tk
from tkinter import messagebox, scrolledtext
import requests
import json
import re
import win32com.client
import pythoncom
import time
import os
import subprocess
import sys
import threading

API_KEY = "YOUR_API_KEY"
GEMINI_URL = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={API_KEY}"

# Chemin du template SolidWorks
SOLIDWORKS_TEMPLATE = r"C:\ProgramData\SOLIDWORKS\SOLIDWORKS 2023\templates\Pièce.PRTDOT"

def call_gemini(prompt_text):
    """Appel à l'API Gemini pour générer du code"""
    headers = {"Content-Type": "application/json"}
    data = {
        "contents": [{"parts": [{"text": prompt_text}]}]
    }
    try:
        resp = requests.post(GEMINI_URL, headers=headers, data=json.dumps(data))
        resp.raise_for_status()
        response = resp.json()
        return response["candidates"][0]["content"]["parts"][0]["text"]
    except Exception as e:
        raise Exception(f"Erreur API Gemini: {str(e)}")

def clean_code(code):
    """Nettoyer le code généré"""
    code = re.sub(r'```python|```py|```', '', code).strip()
    return code

def generate_python_code():
    """Générer du code Python avec Gemini"""
    prompt = text_prompt.get("1.0", tk.END).strip()
    if not prompt:
        messagebox.showwarning("Erreur", "Veuillez entrer un prompt.")
        return

    try:
        # Prompt pour Gemini - création de pièce SolidWorks
        python_prompt = f"""
        Tu es un expert en automation SolidWorks avec Python. Génère du code Python qui utilise l'API COM de SolidWorks via win32com.client.

        Le code doit:
        1. Se connecter à SolidWorks
        2. Créer une nouvelle pièce
        3. Créer une esquisse avec une forme simple
        4. Extruder l'esquisse
        5. Afficher un message de confirmation

        Utilise uniquement ces méthodes éprouvées:
        - win32com.client.Dispatch("SldWorks.Application")
        - swApp.NewDocument(template_path, 0, 0, 0)
        - swApp.ActiveDoc
        - swModel.Extension.SelectByID2()
        - swModel.SketchManager.InsertSketch()
        - swModel.SketchManager.CreateCornerRectangle()
        - swModel.FeatureManager.FeatureExtrusion2()
        - swApp.SendMsgToUser()

        Template à utiliser: {SOLIDWORKS_TEMPLATE}

        Code exemple:
        import win32com.client
        import pythoncom
        import time

        def main():
            try:
                pythoncom.CoInitialize()
                sw_app = win32com.client.Dispatch("SldWorks.Application")
                sw_app.Visible = True
                time.sleep(2)
                
                # Créer nouvelle pièce
                template_path = r"{SOLIDWORKS_TEMPLATE}"
                part_doc = sw_app.NewDocument(template_path, 0, 0, 0)
                time.sleep(1)
                
                sw_model = sw_app.ActiveDoc
                if sw_model is None:
                    raise Exception("No active document")
                
                # Sélectionner plan devant
                sw_model.Extension.SelectByID2("Plan devant", "PLANE", 0, 0, 0, False, 0, None, 0)
                time.sleep(0.5)
                
                # Créer esquisse
                sw_model.SketchManager.InsertSketch(True)
                time.sleep(0.5)
                
                # Dessiner rectangle
                sw_model.SketchManager.CreateCornerRectangle(0, 0, 0, 0.1, 0.05, 0)
                time.sleep(0.5)
                
                # Sortir esquisse
                sw_model.SketchManager.InsertSketch(True)
                time.sleep(0.5)
                
                # Sélectionner esquisse
                sw_model.Extension.SelectByID2("Esquisse1", "SKETCH", 0, 0, 0, False, 0, None, 0)
                time.sleep(0.5)
                
                # Extruder
                feature = sw_model.FeatureManager.FeatureExtrusion2(
                    True, False, False, 0, 0, 0.01, 0.01, 
                    False, False, False, False, 0, 0, 
                    False, False, False, False, True, True, True, 0, 0, False
                )
                
                sw_app.SendMsgToUser("Pièce créée avec succès!")
                
            except Exception as e:
                print(f"Erreur: {{e}}")
                import traceback
                traceback.print_exc()
            finally:
                try:
                    pythoncom.CoUninitialize()
                except:
                    pass

        if __name__ == "__main__":
            main()

        Tâche à réaliser,sans explication et sans SelectByID2 ,juste le code : {prompt}
        """
        
        status_bar.config(text="Appel à Gemini en cours...")
        root.update()
        
        # Appel à l'API Gemini
        generated = call_gemini(python_prompt)
        code = clean_code(generated)
        
        text_code.delete("1.0", tk.END)
        text_code.insert("1.0", code)
        status_bar.config(text="Code généré avec succès via Gemini")
        
    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur avec Gemini: {str(e)}")
        status_bar.config(text="Erreur lors de la génération")

def execute_automatically():
    """Exécution automatique: Génère avec Gemini et exécute"""
    auto_exec_btn.config(state=tk.DISABLED, text="Génération et exécution...")
    status_bar.config(text="Appel à Gemini et exécution en cours...")
    root.update()
    
    def auto_execute_thread():
        try:
            # Étape 1: Générer le code avec Gemini
            generate_python_code()
            time.sleep(1)
            
            # Étape 2: Exécuter le code généré
            code = text_code.get("1.0", tk.END).strip()
            if not code:
                raise Exception("Aucun code généré par Gemini")
            
            status_bar.config(text="Exécution du code Gemini...")
            root.update()
            
            temp_dir = os.environ.get('TEMP', os.getcwd())
            temp_file = os.path.join(temp_dir, "gemini_solidworks.py")
            
            with open(temp_file, 'w', encoding='utf-8') as f:
                f.write(code)
            
            process = subprocess.Popen([sys.executable, temp_file], 
                                     stdout=subprocess.PIPE, 
                                     stderr=subprocess.PIPE,
                                     text=True,
                                     cwd=temp_dir)
            
            stdout, stderr = process.communicate()
            
            result_message = f"Sortie: {stdout}" if stdout else "Aucune sortie"
            if stderr:
                result_message += f"\nErreurs: {stderr}"
            
            if process.returncode == 0:
                root.after(0, lambda: status_bar.config(text="Exécution réussis"))
            else:
                root.after(0, lambda: messagebox.showerror("Erreur", f"Erreur lors de l'exécution:\n{result_message}"))
                root.after(0, lambda: status_bar.config(text="Erreur d'exécution"))
                
        except Exception as e:
            error_msg = f"Erreur: {str(e)}"
            root.after(0, lambda: messagebox.showerror("Erreur", error_msg))
            root.after(0, lambda: status_bar.config(text=error_msg))
        finally:
            root.after(0, lambda: auto_exec_btn.config(state=tk.NORMAL, text="Exécution"))
    
    thread = threading.Thread(target=auto_execute_thread)
    thread.daemon = True
    thread.start()

def execute_create_simple_part():
    """Créer une pièce simple prédéfinie"""
    simple_part_code = f'''import win32com.client
import pythoncom
import time
import traceback

def main():
    try:
        print("Initialisation de COM...")
        pythoncom.CoInitialize()
        
        print("Connexion à SolidWorks via Gemini...")
        sw_app = win32com.client.Dispatch("SldWorks.Application")
        sw_app.Visible = True
        time.sleep(3)
        
        print("Création d'une nouvelle pièce...")
        template_path = r"{SOLIDWORKS_TEMPLATE}"
        part_doc = sw_app.NewDocument(template_path, 0, 0, 0)
        time.sleep(2)
        
        sw_model = sw_app.ActiveDoc
        if sw_model is None:
            raise Exception("Aucun document actif")
        
        print("Sélection du plan devant...")
        sw_model.Extension.SelectByID2("Plan devant", "PLANE", 0, 0, 0, False, 0, None, 0)
        time.sleep(1)
        
        print("Création de l'esquisse...")
        sw_model.SketchManager.InsertSketch(True)
        time.sleep(1)
        
        print("Dessin du rectangle...")
        sw_model.SketchManager.CreateCornerRectangle(0, 0, 0, 0.1, 0.05, 0)
        time.sleep(1)
        
        print("Sortie de l'esquisse...")
        sw_model.SketchManager.InsertSketch(True)
        time.sleep(1)
        
        print("Sélection de l'esquisse...")
        sw_model.Extension.SelectByID2("Esquisse1", "SKETCH", 0, 0, 0, False, 0, None, 0)
        time.sleep(1)
        
        print("Extrusion...")
        feature = sw_model.FeatureManager.FeatureExtrusion2(
            True, False, False, 0, 0, 0.02, 0.02, 
            False, False, False, False, 0, 0, 
            False, False, False, False, True, True, True, 0, 0, False
        )
        
        print("Envoi du message...")
        sw_app.SendMsgToUser("Pièce créée par Gemini")
        
        print("Succès! Pièce créée.")
        return True
        
    except Exception as e:
        print(f"Erreur: {{e}}")
        traceback.print_exc()
        return False
    finally:
        try:
            pythoncom.CoUninitialize()
        except:
            pass

if __name__ == "__main__":
    success = main()
    exit(0 if success else 1)
'''
    
    text_code.delete("1.0", tk.END)
    text_code.insert("1.0", simple_part_code)
    status_bar.config(text="Code de pièce simple généré")
    execute_automatically()

def test_gemini_connection():
    """Tester la connexion à Gemini"""
    try:
        status_bar.config(text="Test connexion Gemini...")
        root.update()
        
        test_prompt = "Génère un message simple: 'Connexion Gemini réussie' en Python"
        response = call_gemini(test_prompt)
        
        messagebox.showinfo("Succès", f"Connexion Gemini OK!\n\nRéponse: {response}")
        status_bar.config(text="Connexion Gemini réussie")
        
    except Exception as e:
        messagebox.showerror("Erreur", f"Échec connexion Gemini: {str(e)}")
        status_bar.config(text="Erreur connexion Gemini")

def save_python_file():
    code = text_code.get("1.0", tk.END).strip()
    if not code:
        messagebox.showwarning("Erreur", "Aucun code à sauvegarder.")
        return
    
    try:
        desktop = os.path.join(os.environ['USERPROFILE'], 'Desktop')
        file_path = os.path.join(desktop, "gemini_solidworks_script.py")
        
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(code)
        
        messagebox.showinfo("Succès", f"Code Gemini sauvegardé: {file_path}")
        status_bar.config(text=f"Code sauvegardé: {file_path}")
        
    except Exception as e:
        error_msg = f"Erreur sauvegarde: {str(e)}"
        messagebox.showerror("Erreur", error_msg)
        status_bar.config(text=error_msg)

def copy_to_clipboard():
    code = text_code.get("1.0", tk.END).strip()
    if not code:
        messagebox.showwarning("Erreur", "Aucun code à copier.")
        return
    
    root.clipboard_clear()
    root.clipboard_append(code)
    status_bar.config(text="Code copié dans le presse-papiers")
    messagebox.showinfo("Succès", "Code Gemini copié!")

# Interface Tkinter
root = tk.Tk()
root.title("AI-Solidworks (Mohamed Hamed)")
root.geometry("500x700")

# Configure grid
root.columnconfigure(0, weight=1)
root.rowconfigure(4, weight=1)

# Prompt section
tk.Label(root, text="prompt:").grid(row=0, column=0, sticky="w", padx=10, pady=(10, 0))
text_prompt = scrolledtext.ScrolledText(root, height=4, width=100)
text_prompt.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 10))
text_prompt.insert("1.0", "Créer une pièce avec un rectangle extrudé de 20mm")

# Generated code section
tk.Label(root, text="").grid(row=2, column=0, sticky="w", padx=10)
text_code = scrolledtext.ScrolledText(root, height=18, width=100, font=("Courier New", 10))
text_code.grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 10))

# Code par défaut
default_code = f""""""
text_code.insert("1.0", default_code)

# Boutons Gemini
gemini_frame = tk.Frame(root)
gemini_frame.grid(row=4, column=0, pady=10)
# Bouton principal
auto_exec_btn = tk.Button(root, text="EXÉCUTION", command=execute_automatically, 
                         bg="#4CAF50", fg="white", font=("Arial", 12, "bold"), height=2)
auto_exec_btn.grid(row=5, column=0, pady=10, padx=10, sticky="ew")

# Bouton pièce simple
simple_part_btn = tk.Button(root, text="CRÉER PIÈCE SIMPLE", command=execute_create_simple_part, 
                          bg="#FF9800", fg="white", font=("Arial", 10, "bold"), height=1)
simple_part_btn.grid(row=6, column=0, pady=5, padx=10, sticky="ew")

# Buttons frame
btn_frame = tk.Frame(root)
btn_frame.grid(row=7, column=0, pady=5)

save_btn = tk.Button(btn_frame, text="Sauvegarder", command=save_python_file, bg="#607D8B", fg="white", width=12)
save_btn.pack(side=tk.LEFT, padx=5)

copy_btn = tk.Button(btn_frame, text="Copier", command=copy_to_clipboard, bg="#795548", fg="white", width=10)
copy_btn.pack(side=tk.LEFT, padx=5)

# Instructions
instructions = f"""
Mohamed Hamed 2025
"""
instructions_label = tk.Label(root, text=instructions, justify=tk.LEFT, foreground="gray")
instructions_label.grid(row=8, column=0, sticky="w", padx=10, pady=10)

# Status bar
status_bar = tk.Label(root, text="Prêt", bd=1, relief=tk.SUNKEN, anchor=tk.W)
status_bar.grid(row=9, column=0, sticky="ew")

root.mainloop()