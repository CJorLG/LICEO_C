import pandas as pd
from unidecode import unidecode

# ---------------------------------------------------------
# Función para normalizar nombres
# ---------------------------------------------------------



def normalizar_nombre(nombre):
    if pd.isna(nombre):
        return ""

    # Quita acentos
    nombre = unidecode(nombre)

    # Mayúsculas
    nombre = nombre.upper()

    # Quita espacios extra
    nombre = " ".join(nombre.split())

    return nombre


# ---------------------------------------------------------
# Correge orden de nombres usando lista base
# Funciona para AP1 AP2 NOMBRE (3 palabras)
# ---------------------------------------------------------

def corregir_orden(nombre, lista_base):
    """
    Para corregir el orden del nombre aunque venga invertido o mezclado.
    Aquí usa la lista base como referencia para identificar al alumno.
    """

    palabras = nombre.split()

    # Si no tiene al menos 3 palabras (ej. AP1 AP2 NOMBRE), lo dejamos igual
    if len(palabras) < 3:
        return nombre

    # Lo convierte en un set para comparar sin importar el orden
    set_nuevo = set(palabras)

    # Revisa todos los nombres de la lista base
    for nombre_base in lista_base:
        palabras_base = nombre_base.split()
        set_base = set(palabras_base)

        # Aquí evalua si contienen exactamente las mismas palabras → mismo alumno
        if set_nuevo == set_base:
            return nombre_base   # Regresa el formato correcto de la base

    # Si no coincide con nadie en la base, se deja como está
    return nombre


# ---------------------------------------------------------
# Función principal
# ---------------------------------------------------------

def actualizar_creditos(archivo_base, archivo_nuevo, archivo_salida):
    # Leer archivos Excel
    base = pd.read_excel(archivo_base)
    nuevo = pd.read_excel(archivo_nuevo)

    # Normalizar nombres
    base["Nombre"] = base["Nombre"].apply(normalizar_nombre)
    nuevo["Nombre"] = nuevo["Nombre"].apply(normalizar_nombre)

    # Crear conjunto de nombres de la base (normalizados)
    lista_base = set(base["Nombre"])

    # Corregir orden en el nuevo archivo
    nuevo["Nombre"] = nuevo["Nombre"].apply(lambda x: corregir_orden(x, lista_base))

    # Convertir a diccionarios {nombre: creditos}
    dict_base = dict(zip(base["Nombre"], base["Creditos"]))
    dict_nuevo = dict(zip(nuevo["Nombre"], nuevo["Creditos"]))

    # Diccionario final
    resultado = {}

    # --------------------------------------------
    # A) Sumar créditos de alumnos existentes
    # --------------------------------------------
    for alumno in dict_base:
        if alumno in dict_nuevo:
            resultado[alumno] = dict_base[alumno] + dict_nuevo[alumno]
        # Si ya no está en la nueva lista, NO se agrega (baja)

    # --------------------------------------------
    # B) Agregar alumnos nuevos
    # --------------------------------------------
    for alumno in dict_nuevo:
        if alumno not in resultado:
            resultado[alumno] = dict_nuevo[alumno]

    # Convertir a DataFrame final
    df_final = pd.DataFrame({
        "Nombre": list(resultado.keys()),
        "Creditos": list(resultado.values())
    })

    # Ordenar alfabéticamente
    df_final = df_final.sort_values(by="Nombre")

    # Guardar archivo resultante
    df_final.to_excel(archivo_salida, index=False)
    print(f"Archivo actualizado guardado como: {archivo_salida}")


# ---------------------------------------------------------
# Punto de ejecución
# ---------------------------------------------------------

if __name__ == "__main__":
    actualizar_creditos(
        archivo_base=r"C:\Users\jorel\Desktop\LICEO_C\Files\Documents\pruebaLectura.xlsx",
        archivo_nuevo=r"C:\Users\jorel\Desktop\LICEO_C\Files\Documents\pruebaLectura2.xlsx",
        archivo_salida=r"C:\Users\jorel\Desktop\LICEO_C\Files\Documents\creditos_final.xlsx",
    )