#contador_palabras
#py -3.13 D:/#contador_palabras.py D:\Juan_Sierra_IA.docx
import docx
import re
from collections import Counter
import string
import matplotlib.pyplot as plt
from wordcloud import WordCloud

# Lista extendida de stopwords en español
stopwords_es = {
    'el', 'la', 'los', 'las', 'un', 'una', 'unos', 'unas',
    'de', 'del', 'al', 'y', 'o', 'u', 'que', 'como', 'para',
    'por', 'con', 'sin', 'sobre', 'entre', 'hasta', 'desde',
    'a', 'en', 'es', 'son', 'ser', 'fue', 'era', 'soy', 'eres',
    'está', 'están', 'estaba', 'estaban', 'muy', 'más', 'menos',
    'también', 'pero', 'aunque', 'si', 'sí', 'no', 'ya', 'lo',
    'le', 'se', 'me', 'te', 'nos', 'os', 'su', 'sus', 'mi', 'mis',
    'tu', 'tus', 'nuestro', 'nuestra', 'vuestro', 'vuestra',
    'este', 'esta', 'estos', 'estas', 'ese', 'esa', 'esos', 'esas',
    'aquí', 'allí', 'ahí', 'entonces', 'cuando', 'donde',
    'porque', 'mientras', 'pues', 'además', 'incluso'
}

def extraer_texto_docx(ruta_archivo):
    try:
        doc = docx.Document(ruta_archivo)
        texto_completo = [p.text for p in doc.paragraphs if p.text.strip()]
        return '\n'.join(texto_completo)
    except Exception as e:
        print(f"Error al leer el archivo: {e}")
        return ""

def limpiar_y_dividir_palabras(texto, incluir_numeros=False):
    texto = texto.lower()
    if incluir_numeros:
        texto_limpio = re.sub(r'[^\w\sáéíóúüñ]', ' ', texto)
        palabras = [p for p in texto_limpio.split() if p.strip()]
    else:
        palabras = re.findall(r'\b[a-záéíóúüñ]+\b', texto)
    return palabras

def contar_palabras_docx(ruta_archivo, incluir_numeros=False, palabras_minimas=1, excluir_stopwords=True):
    texto = extraer_texto_docx(ruta_archivo)
    if not texto:
        print("No se pudo extraer texto del documento.")
        return Counter()
    
    palabras = limpiar_y_dividir_palabras(texto, incluir_numeros)
    palabras_filtradas = [p for p in palabras if len(p) >= palabras_minimas]

    if excluir_stopwords:
        palabras_filtradas = [p for p in palabras_filtradas if p not in stopwords_es]

    return Counter(palabras_filtradas)

def mostrar_estadisticas(contador, top_n=20):
    if not contador:
        print("No hay palabras para analizar.")
        return
    
    total = sum(contador.values())
    print(f"\n{'='*50}")
    print(f"ESTADÍSTICAS DEL DOCUMENTO")
    print(f"{'='*50}")
    print(f"Total de palabras: {total}")
    print(f"Palabras únicas: {len(contador)}")
    print(f"\nTOP {top_n} PALABRAS MÁS FRECUENTES")
    print(f"{'Palabra':<20} {'Frecuencia':<10} {'Porcentaje'}")
    print("-" * 50)
    
    for palabra, frecuencia in contador.most_common(top_n):
        porcentaje = (frecuencia / total) * 100
        print(f"{palabra:<20} {frecuencia:<10} {porcentaje:.2f}%")

def crear_grafico_barras(contador, top_n=25, archivo_imagen="grafico_palabras.png"):
    if not contador:
        print("No hay datos para graficar.")
        return
    
    palabras_top = contador.most_common(top_n)
    palabras = [item[0] for item in palabras_top]
    frecuencias = [item[1] for item in palabras_top]
    
    plt.figure(figsize=(15, 10))
    barras = plt.bar(range(len(palabras)), frecuencias, color='skyblue', alpha=0.8)
    
    plt.title(f'Top {top_n} Palabras Más Frecuentes', fontsize=16, fontweight='bold', pad=20)
    plt.xlabel('Palabras', fontsize=12, fontweight='bold')
    plt.ylabel('Frecuencia', fontsize=12, fontweight='bold')
    plt.xticks(range(len(palabras)), palabras, rotation=45, ha='right')
    
    for i, barra in enumerate(barras):
        altura = barra.get_height()
        plt.text(barra.get_x() + barra.get_width()/2., altura + 0.1,
                 f'{int(altura)}', ha='center', va='bottom', fontsize=9)
    
    plt.tight_layout()
    plt.grid(axis='y', alpha=0.3, linestyle='--')
    
    try:
        plt.savefig(archivo_imagen, dpi=300, bbox_inches='tight')
        print(f"\nGráfico guardado como: {archivo_imagen}")
    except Exception as e:
        print(f"Error al guardar el gráfico: {e}")
    
    try:
        plt.show()
    except Exception as e:
        print(f"No se puede mostrar el gráfico en este entorno: {e}")

def crear_nube_palabras(contador, top_n=50, archivo_imagen="nube_palabras.png"):
    if not contador:
        print("No hay datos para la nube de palabras.")
        return
    
    palabras_top = dict(contador.most_common(top_n))
    
    try:
        nube = WordCloud(width=800, height=600, background_color='white',
                         colormap='viridis', max_words=top_n).generate_from_frequencies(palabras_top)
        
        plt.figure(figsize=(10, 8))
        plt.imshow(nube, interpolation='bilinear')
        plt.axis("off")
        plt.title(f"Nube de Palabras (Top {top_n})", fontsize=16, fontweight='bold', pad=20)
        plt.tight_layout()
        
        plt.savefig(archivo_imagen, dpi=300, bbox_inches='tight')
        plt.show()
        
        print(f"\nNube de palabras guardada como: {archivo_imagen}")
    
    except Exception as e:
        print(f"Error al crear la nube de palabras: {e}")

def guardar_resultados(contador, archivo_salida="conteo_palabras.txt"):
    try:
        with open(archivo_salida, 'w', encoding='utf-8') as f:
            f.write("CONTEO COMPLETO DE PALABRAS\n")
            f.write("=" * 50 + "\n\n")
            total = sum(contador.values())
            f.write(f"Total de palabras: {total}\n")
            f.write(f"Palabras únicas: {len(contador)}\n\n")
            for palabra, frecuencia in contador.most_common():
                porcentaje = (frecuencia / total) * 100
                f.write(f"{palabra}: {frecuencia} ({porcentaje:.2f}%)\n")
        print(f"\nResultados guardados en: {archivo_salida}")
    except Exception as e:
        print(f"Error al guardar archivo: {e}")

def main():
    print("CONTADOR DE PALABRAS EN DOCUMENTOS DOCX")
    print("=" * 50)
    
    ruta_archivo = input("\nIngresa la ruta del archivo .docx: ").strip()
    if not ruta_archivo.endswith('.docx'):
        print("Advertencia: El archivo no tiene extensión .docx")
    
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
        top_n = int(input("¿Cuántas palabras más frecuentes mostrar? (default: 50): ") or "50")
    except ValueError:
        top_n = 50

    print("\nProcesando documento...")
    contador = contar_palabras_docx(ruta_archivo, incluir_numeros, palabras_minimas, excluir_stopwords=True)

    mostrar_estadisticas(contador, top_n)
    crear_grafico_barras(contador, top_n=top_n, archivo_imagen="grafico_palabras.png")
    crear_nube_palabras(contador, top_n=top_n, archivo_imagen="nube_palabras.png")
    
    guardar = input("\n¿Deseas guardar todos los resultados en un archivo? (s/n): ").lower() == 's'
    if guardar:
        nombre_archivo = input("Nombre del archivo (default: conteo_palabras.txt): ").strip()
        if not nombre_archivo:
            nombre_archivo = "conteo_palabras.txt"
        guardar_resultados(contador, nombre_archivo)

if __name__ == "__main__":
    main()
