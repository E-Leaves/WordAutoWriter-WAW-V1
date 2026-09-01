# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
import threading
import time
import win32com.client


pause = False
stop = False



def start_writer():

    global pause, stop

    pause = False
    stop = False


    titre = title_entry.get().strip()

    texte = text_box.get("1.0", tk.END).strip()


    if texte == "":
        messagebox.showwarning(
            "Attention",
            "Veuillez entrer un texte"
        )
        return


    vitesse = int(speed_var.get())


    chemin = filedialog.asksaveasfilename(

        title="Enregistrer le document",

        defaultextension=".docx",

        filetypes=[
            ("Word Document","*.docx")
        ]

    )


    if not chemin:
        return



    word = win32com.client.Dispatch(
        "Word.Application"
    )


    word.Visible = True


    document = word.Documents.Add()


    curseur = word.Selection



    # Titre Word

    curseur.Font.Bold = True

    curseur.Font.Size = 18


    curseur.TypeText(titre)


    curseur.TypeParagraph()



    # Developer
    curseur.Font.Bold = False
    curseur.Font.Size = 12

    curseur.TypeText("Developed by: Mohammed KADRI")
    curseur.TypeParagraph()

    # Date
    date = datetime.now().strftime("%d/%m/%Y")
    curseur.TypeText("Date : " + date)

    curseur.TypeParagraph()
    curseur.TypeParagraph()



    # Texte lettre par lettre

    for lettre in texte:


        if stop:
            break


        while pause:

            time.sleep(0.1)



        curseur.TypeText(lettre)

        time.sleep(
            vitesse / 1000
        )



    document.SaveAs2(chemin)



    messagebox.showinfo(
        "Terminé",
        "Document Word enregistré"
    )





def lancer_thread():

    t = threading.Thread(
        target=start_writer
    )

    t.start()




def pause_resume():

    global pause

    pause = not pause


    if pause:

        pause_button.config(
            text="Reprendre"
        )

    else:

        pause_button.config(
            text="Pause"
        )




def stop_writer():

    global stop

    stop = True




def clear_text():

    text_box.delete(
        "1.0",
        tk.END
    )





# ---------------- INTERFACE ----------------


app = tk.Tk()

app.title(
    "Word Auto Writer DAW V1"
)


# taille selon écran

largeur = app.winfo_screenwidth()

hauteur = app.winfo_screenheight()


app.geometry(
    f"{largeur}x{hauteur}"
)


app.configure(
    bg="#808080"
)



frame = tk.Frame(

    app,

    bg="#808080",

    padx=30,

    pady=30

)


frame.pack(

    expand=True,

    fill="both"

)



# Titre


tk.Label(

    frame,

    text="Titre du document",

    font=("Arial",14,"bold"),

    bg="#808080"

).pack()



title_entry = tk.Entry(

    frame,

    font=("Arial",14)

)

title_entry.pack(

    fill="x",

    pady=10

)



# Texte


tk.Label(

    frame,

    text="Texte à écrire",

    font=("Arial",14,"bold"),

    bg="#808080"

).pack()



text_box = tk.Text(

    frame,

    height=15,

    font=("Arial",13)

)


text_box.pack(

    fill="both",

    expand=True

)



# vitesse


tk.Label(

    frame,

    text="Vitesse",

    font=("Arial",14,"bold"),

    bg="#808080"

).pack()



speed_var = tk.StringVar()

speed_var.set("100")



tk.OptionMenu(

    frame,

    speed_var,

    "50",

    "100",

    "200",

    "300"

).pack()



# boutons



tk.Button(

    frame,

    text="Commencer",

    bg="#2563eb",

    fg="white",

    font=("Arial",12),

    command=lancer_thread

).pack(

    fill="x",

    pady=5

)



pause_button = tk.Button(

    frame,

    text="Pause",

    command=pause_resume

)


pause_button.pack(

    fill="x"

)



tk.Button(

    frame,

    text="Arrêter",

    bg="red",

    fg="white",

    command=stop_writer

).pack(

    fill="x",

    pady=5

)



tk.Button(

    frame,

    text="Effacer",

    command=clear_text

).pack(

    fill="x"

)



app.mainloop()