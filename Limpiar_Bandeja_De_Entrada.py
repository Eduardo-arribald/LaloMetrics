import pandas as pd
import win32com.client as wincl

# Escribe la cuenta que vamos a limpiar
accountName = "TU CORREO"

# Escribe el nombre de la carpeta que vamos a limpiar
folderToClean = "Bandeja de entrada"


# Inicializamos aplicación de outlook
ns = wincl.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Buscamos la cuenta en cuestión
for store in ns.Stores:
    if store.DisplayName == accountName:
        print(store.DisplayName)
        mainStore = store

# Asignamos la carpeta que limpiaremos a una variable inbox
inbox = mainStore.GetRootFolder().Folders[folderToClean]

# Creamos una lista vacía para almacenar los senders
sendersList = []

for mail in inbox.Items:
    sendersList.append(mail.SenderEmailAddress)

# Guardamos los senders en un diccionario
sendersDict = {"Senders": sendersList}

# Convertimos el diccionario a un dataframe
sendersDF = pd.DataFrame(sendersDict)

# El dataframe lo convertimos en una tabla pivote con el conteo de correos por sender
sendersPT = pd.crosstab(
    sendersDF["Senders"],
    columns = "count"
)

# Organizamos esta tabla pivote de mayor a menor, para ver aquellos senders con mayor volumen de correos en nuestra carpeta
sendersPT = sendersPT.sort_values(
    by = "count",
    ascending = False
)

# Vemos los primeros n senders con mayor volumen de correos
n = 5   # Puedes cambiarlo por el número de senders que desees ver
print(sendersPT.head(n))