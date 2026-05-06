import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import requests
import os

# ===============================
# CONFIGURATION
# ===============================

DATABRICKS_URL = "https://adb-811328758478862.2.azuredatabricks.net/serving-endpoints/databricks-claude-sonnet-4-6/invocations"

DATABRICKS_TOKEN = os.getenv("DATABRICKS_TOKEN")

if not DATABRICKS_TOKEN:
    raise RuntimeError("DATABRICKS_TOKEN is not set")


HEADERS = {
    "Authorization": f"Bearer {DATABRICKS_TOKEN}",
    "Content-Type": "application/json"
}

# ===============================
# AI CALL FUNCTION
# ===============================
def reframe_problem():
    user_text = input_text.get("1.0", tk.END).strip()

    if not user_text:
        messagebox.showwarning("Input Missing", "Please enter a problem or load a file.")
        return

    payload = {
        "messages": [
            {
                "role": "system",
                "content": "You are an assistant that rewrites and elaborates problem statements clearly and professionally."
            },
            {
                "role": "user",
                "content": f"Reframe and elaborate the following problem:\n\n{user_text}"
            }
        ],
        "temperature": 0.3,
        "max_tokens": 400
    }

    try:
        response = requests.post(
            DATABRICKS_URL,
            headers=HEADERS,
            json=payload,
            timeout=60
        )
        response.raise_for_status()

        result = response.json()
        ai_output = result["choices"][0]["message"]["content"]

        output_text.delete("1.0", tk.END)
        output_text.insert(tk.END, ai_output)

    except Exception as e:
        messagebox.showerror("Databricks Error", str(e))


# ===============================
# FILE BROWSE FUNCTION
# ===============================
def browse_file():
    file_path = filedialog.askopenfilename(
        filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")]
    )

    if file_path:
        try:
            with open(file_path, "r", encoding="utf-8") as file:
                content = file.read()

            input_text.delete("1.0", tk.END)
            input_text.insert(tk.END, content)

        except Exception as e:
            messagebox.showerror("File Error", str(e))


# ===============================
# UI SETUP
# ===============================
root = tk.Tk()
root.title("Problem Reframing Tool")
root.geometry("900x700")

tk.Label(root, text="Box 1: User Problem / File Content").pack(anchor="w", padx=10)
input_text = scrolledtext.ScrolledText(root, height=10)
input_text.pack(fill="both", padx=10, pady=5)

button_frame = tk.Frame(root)
button_frame.pack(pady=10)

tk.Button(button_frame, text="Browse File", width=20, command=browse_file).pack(side="left", padx=10)
tk.Button(button_frame, text="Reframe Problem", width=20, command=reframe_problem).pack(side="left", padx=10)

tk.Label(root, text="Box 3: AI Reframed & Elaborated Problem").pack(anchor="w", padx=10)
output_text = scrolledtext.ScrolledText(root, height=15)
output_text.pack(fill="both", padx=10, pady=5)

root.mainloop()