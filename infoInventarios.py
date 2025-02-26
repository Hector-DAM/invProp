import pandas as pd

# Cargar la tabla de UPC
tabla_upc_root = "C:/DATA/Catalogo/TABLA UPC.xlsx"
tabla_upc = pd.read_excel(tabla_upc_root)
tabla_upc['UPC'] = tabla_upc['UPC'].astype(str)
tabla_upc['UPC'] = tabla_upc['UPC'].str.replace(".0", "")

# Cargar el inventario de la semana
inventario_root = "C:/DATA/Codigos/PropuestaInventarios/Inventario SMY XRS 2_20_2025 6_50.xlsx"
inventario = pd.read_excel(inventario_root)
inventario["UPC"] = inventario["UPC"].astype(str)
inventario["UPC"] = inventario["UPC"].str.replace(".0", "")

# Cargar la información de las tiendas
tiendas_root = "C:/DATA/Codigos/PropuestaInventarios/Tiendas M3.xlsx"
tiendas = pd.read_excel(tiendas_root)

# Filtrar solo el inventario de tiendas
inventario = inventario[inventario["WH"] == "XRS"]

# Hacer un merge con la tabla de UPC
inventario = pd.merge(inventario, tabla_upc, how="left", on="UPC")

# Hacer un merge con la tabla de tiendas
inventarioFinal = pd.merge(inventario, tiendas, how="left", on="STORE")

# Crear la columna BARCODE y EstiloColor
inventarioFinal["BARCODE"] = inventarioFinal["STYLE M3"] + inventarioFinal["Color Code"]
inventarioFinal["EstiloColor"] = inventarioFinal["STYLE_y"] + "-" + inventarioFinal["Color Name"]

# Filtrar la marca CALZANETTO
inventarioFinal = inventarioFinal[inventarioFinal["Brand"] != "CALZANETTO"]

# Seleccionar las columnas necesarias y ordenar
inventarioFinal = inventarioFinal[["BARCODE", "Tienda", "UPC", "EstiloColor", "Size", "Brand", "AVAILABLE"]]
inventarioFinal = inventarioFinal.sort_values(by=["BARCODE", "Tienda"])

# Seleccionar el 25% superior del inventario ordenado
percent = 0.25
num_registros = int(len(inventarioFinal) * percent)
muestra = inventarioFinal.head(num_registros)

# Crear un archivo Excel con una hoja por tienda (propuesta agrupada)
with pd.ExcelWriter("PropuestaConteo_PorTienda_Ajustada.xlsx") as writer:
    for tienda in muestra["Tienda"].unique():  # Iterar sobre cada tienda única
        # Filtrar los datos para la tienda actual
        df_tienda = muestra[muestra["Tienda"] == tienda]
        
        # Agrupar por Tienda, BARCODE, EstiloColor y Brand, y pivotar las tallas
        df_propuesta = df_tienda.pivot_table(
            index=["Tienda", "BARCODE", "EstiloColor", "Brand"],
            columns="Size",
            values="AVAILABLE",
            fill_value=0
        ).reset_index()
        
        # Renombrar las columnas de tallas
        df_propuesta.columns = [f"Talla{col}" if str(col).isdigit() else col for col in df_propuesta.columns]
        
        # Guardar la propuesta en una hoja con el nombre de la tienda
        df_propuesta.to_excel(writer, sheet_name=tienda, index=False)

print("Propuesta de conteo agrupada generada y guardada en 'PropuestaConteo_PorTienda_Ajustada.xlsx'")

# Crear un archivo Excel con solo los UPC (propuesta de UPC)
upc_lista = muestra[["UPC"]].drop_duplicates()  # Seleccionar solo la columna UPC y eliminar duplicados
upc_lista.to_excel("Lista_UPC_Propuesta.xlsx", index=False)

print("Lista de UPC generada y guardada en 'Lista_UPC_Propuesta.xlsx'")

# Crear un archivo Excel con una hoja por tienda (propuesta de UPC por tienda con cantidades)
with pd.ExcelWriter("Lista_UPC_PorTienda_ConCantidades.xlsx") as writer:
    for tienda in muestra["Tienda"].unique():  # Iterar sobre cada tienda única
        # Filtrar los datos para la tienda actual
        df_tienda = muestra[muestra["Tienda"] == tienda]
        
        # Agrupar por UPC y sumar las cantidades (AVAILABLE)
        upc_tienda = df_tienda.groupby("UPC", as_index=False)["AVAILABLE"].sum()
        
        # Guardar la lista de UPC con cantidades en una hoja con el nombre de la tienda
        upc_tienda.to_excel(writer, sheet_name=tienda, index=False)

print("Lista de UPC por tienda con cantidades generada y guardada en 'Lista_UPC_PorTienda_ConCantidades.xlsx'")