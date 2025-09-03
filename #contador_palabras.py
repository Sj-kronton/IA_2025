#contador_palabras
#py -3.13 D:/#contador_palabras.py D:\Juan_Sierra_IA.docx
import docx
import re
from collections import Counter
import string

def extraer_texto_docx(ruta_archivo):
    """
    Extrae todo el texto de un archivo .docx
    
    Args:
        ruta_archivo (str): Ruta al archivo .docx
        
    Returns:
        str: Texto completo del documento
    """
    try:
        doc = docx.Document(ruta_archivo)
        texto_completo = []
        
        # Extraer texto de todos los párrafos
        for parrafo in doc.paragraphs:
            if parrafo.text.strip():  # Solo agregar párrafos no vacíos
                texto_completo.append(parrafo.text)
        
        return '\n'.join(texto_completo)
    
    except Exception as e:
        print(f"Error al leer el archivo: {e}")
        return ""

def limpiar_y_dividir_palabras(texto, incluir_numeros=False):
    """
    Limpia el texto y lo divide en palabras
    
    Args:
        texto (str): Texto a procesar
        incluir_numeros (bool): Si incluir números en el conteo
        
    Returns:
        list: Lista de palabras limpias
    """
    # Convertir a minúsculas
    texto = texto.lower()
    
    # Debug: mostrar texto antes de limpiar
    print(f"DEBUG: Texto a procesar (primeros 100 caracteres): {texto[:100]}...")
    
    if incluir_numeros:
        # Remover solo algunos signos de puntuación, conservar números
        texto_limpio = re.sub(r'[^\w\sáéíóúüñ]', ' ', texto)
        palabras = [palabra for palabra in texto_limpio.split() if palabra.strip()]
    else:
        # Solo conservar letras (incluye acentos y ñ)
        palabras = re.findall(r'\b[a-záéíóúüñA-ZÁÉÍÓÚÜÑ]+\b', texto)
    
    # Debug: mostrar resultado
    print(f"DEBUG: Palabras encontradas: {len(palabras)}")
    if palabras:
        print(f"DEBUG: Primeras 10 palabras: {palabras[:10]}")
    else:
        print("DEBUG: No se encontraron palabras")
    
    return palabras

def contar_palabras_docx(ruta_archivo, incluir_numeros=False, palabras_minimas=1):
    """
    Cuenta la frecuencia de palabras en un documento .docx
    
    Args:
        ruta_archivo (str): Ruta al archivo .docx
        incluir_numeros (bool): Si incluir números en el conteo
        palabras_minimas (int): Longitud mínima de palabras a contar
        
    Returns:
        Counter: Objeto Counter con el conteo de palabras
    """
    # Extraer texto del documento
    texto = extraer_texto_docx(ruta_archivo)
    
    if not texto:
        print("No se pudo extraer texto del documento.")
        return Counter()
    
    # Limpiar y dividir en palabras
    palabras = limpiar_y_dividir_palabras(texto, incluir_numeros)
    
    # Filtrar palabras por longitud mínima
    palabras_filtradas = [p for p in palabras if len(p) >= palabras_minimas]
    
    # Contar frecuencias
    contador = Counter(palabras_filtradas)
    
    return contador

def mostrar_estadisticas(contador, top_n=20):
    """
    Muestra estadísticas del conteo de palabras
    
    Args:
        contador (Counter): Contador de palabras
        top_n (int): Número de palabras más frecuentes a mostrar
    """
    if not contador:
        print("No hay palabras para analizar.")
        return
    
    total_palabras = sum(contador.values())
    palabras_unicas = len(contador)
    
    print(f"\n{'='*50}")
    print(f"ESTADÍSTICAS DEL DOCUMENTO")
    print(f"{'='*50}")
    print(f"Total de palabras: {total_palabras}")
    print(f"Palabras únicas: {palabras_unicas}")
    print(f"\n{'='*50}")
    print(f"TOP {top_n} PALABRAS MÁS FRECUENTES")
    print(f"{'='*50}")
    print(f"{'Palabra':<20} {'Frecuencia':<10} {'Porcentaje'}")
    print("-" * 50)
    
    for palabra, frecuencia in contador.most_common(top_n):
        porcentaje = (frecuencia / total_palabras) * 100
        print(f"{palabra:<20} {frecuencia:<10} {porcentaje:.2f}%")

def guardar_resultados(contador, archivo_salida="conteo_palabras.txt"):
    """
    Guarda los resultados en un archivo de texto
    
    Args:
        contador (Counter): Contador de palabras
        archivo_salida (str): Nombre del archivo de salida
    """
    try:
        with open(archivo_salida, 'w', encoding='utf-8') as f:
            f.write("CONTEO COMPLETO DE PALABRAS\n")
            f.write("=" * 50 + "\n\n")
            
            total_palabras = sum(contador.values())
            f.write(f"Total de palabras: {total_palabras}\n")
            f.write(f"Palabras únicas: {len(contador)}\n\n")
            
            for palabra, frecuencia in contador.most_common():
                porcentaje = (frecuencia / total_palabras) * 100
                f.write(f"{palabra}: {frecuencia} ({porcentaje:.2f}%)\n")
        
        print(f"\nResultados guardados en: {archivo_salida}")
        
    except Exception as e:
        print(f"Error al guardar archivo: {e}")

def main():
    """
    Función principal del programa
    """
    print("CONTADOR DE PALABRAS EN DOCUMENTOS DOCX")
    print("=" * 50)
    
    # Solicitar ruta del archivo
    ruta_archivo = input("\nIngresa la ruta del archivo .docx: ").strip()
    
    if not ruta_archivo.endswith('.docx'):
        print("Advertencia: El archivo no tiene extensión .docx")
    
    # Opciones de procesamiento
    print("\nOpciones de procesamiento:")
    incluir_numeros = input("¿Incluir números en el conteo? (s/n): ").lower() == 's'
    
    try:
        palabras_minimas = int(input("Longitud mínima de palabras (default: 1): ") or "1")
        if palabras_minimas < 1:
            palabras_minimas = 1
            print("Se estableció longitud mínima en 1")
    except ValueError:
        palabras_minimas = 1
        print("Valor inválido, se estableció longitud mínima en 1")
    
    try:
        top_n = int(input("¿Cuántas palabras más frecuentes mostrar? (default: 20): ") or "20")
    except ValueError:
        top_n = 20
    
    # Procesar documento
    print("\nProcesando documento...")
    contador = contar_palabras_docx(ruta_archivo, incluir_numeros, palabras_minimas)
    
    # Mostrar resultados
    mostrar_estadisticas(contador, top_n)
    
    # Opción para guardar resultados
    guardar = input("\n¿Deseas guardar todos los resultados en un archivo? (s/n): ").lower() == 's'
    if guardar:
        nombre_archivo = input("Nombre del archivo (default: conteo_palabras.txt): ").strip()
        if not nombre_archivo:
            nombre_archivo = "conteo_palabras.txt"
        guardar_resultados(contador, nombre_archivo)

# Ejemplo de uso alternativo (sin interfaz interactiva)
def ejemplo_uso_directo():
    """
    Ejemplo de uso directo del programa sin interfaz interactiva
    """
    # Cambiar por la ruta de tu archivo
    ruta_archivo = "mi_documento.docx"
    
    # Contar palabras
    contador = contar_palabras_docx(ruta_archivo, incluir_numeros=False, palabras_minimas=2)
    
    # Mostrar estadísticas
    mostrar_estadisticas(contador, top_n=15)
    
    # Guardar resultados
    guardar_resultados(contador, "resultados_conteo.txt")

if __name__ == "__main__":
    main()
    
    # Descomentar la siguiente línea para usar el ejemplo directo
    # ejemplo_uso_directo()