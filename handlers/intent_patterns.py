import re

# Detección de intención de edición: verbos que implican modificar el archivo
_RE_EDICION = re.compile(
    r"\b("
    r"a[ñn]ade[r]?\s|agrega[r]?\s|crea\s+(?:una\s+)?columna|nueva\s+columna|"
    r"calcula\s+(?:una\s+)?columna|"
    r"ordena[r]?\s|ordena\s+(?:el|los|por)|"
    r"elimina[r]?\s+(?:los\s+|las\s+)?duplicados|"
    r"quita[r]?\s+(?:los\s+|las\s+)?duplicados|"
    r"borra[r]?\s+(?:los\s+|las\s+)?duplicados|"
    r"elimina[r]?\s+(?:la\s+|el\s+)?columna|"
    r"quita[r]?\s+(?:la\s+|el\s+)?columna|"
    r"borra[r]?\s+(?:la\s+|el\s+)?columna|"
    r"rellena[r]?\s+(?:los\s+|las\s+)?(?:vac[ií]os?|nulos?|huecos?)|"
    r"completa[r]?\s+(?:los\s+|las\s+)?(?:vac[ií]os?|nulos?)|"
    r"renombra[r]?\s|cambia[r]?\s+el\s+nombre\s+de\s|"
    r"aplica[r]?\s+formato\s+condicional|aplica[r]?\s+color|"
    r"pinta[r]?\s+(?:en\s+)?(?:rojo|verde|amarillo|naranja|azul)|colorea[r]?\s|"
    r"normaliza[r]?\s+(?:el\s+)?(?:texto|datos)|"
    r"limpia[r]?\s+(?:los\s+)?espacios|quita[r]?\s+(?:los\s+)?espacios|"
    r"(?:unifica[r]?|convierte[r]?\s+a|pon[er]?\s+en)\s+(?:may[uú]sculas?|min[uú]sculas?)|"
    r"capitaliza[r]?\s|"
    r"(?:corrige[r]?|estandariza[r]?|formatea[r]?)\s+(?:las\s+)?fechas?|"
    r"(?:des)?pivotea[r]?|"
    r"convierte[r]?\s+(?:las\s+)?columnas?\s+en\s+filas?|"
    r"convierte[r]?\s+(?:las\s+)?filas?\s+en\s+columnas?|"
    r"(?:meses?|trimestres?|periodos?)\s+en\s+filas?|"
    r"reemplaza[r]?\s|sustituye[r]?\s|cambia[r]?\s+(?:todos?\s+los?\s+|todas?\s+las?\s+)?(?:valores?|celdas?|textos?)\s|"
    r"busca[r]?\s+y\s+reemplaza[r]?|"
    r"divide[r]?\s+(?:la\s+)?columna|separa[r]?\s+(?:la\s+)?columna|partir\s+(?:la\s+)?columna|"
    r"concat[eé]na[r]?\s|une[r]?\s+(?:las\s+)?columnas?|junta[r]?\s+(?:las\s+)?columnas?"
    r")\b",
    re.IGNORECASE,
)

# Detección de combinación de dos archivos
_RE_COMBINAR = re.compile(
    r"\b("
    r"une[r]?\s|combina[r]?\s|junta[r]?\s|mezcla[r]?\s|fusiona[r]?\s|"
    r"cruza[r]?\s|cruce\s+(?:con|de)|"
    r"(?:los\s+)?dos\s+archivos|ambos\s+archivos|"
    r"merge[r]?\s|join\s"
    r")\b",
    re.IGNORECASE,
)

# Detección de creación de Excel desde descripción
_RE_CREAR_EXCEL = re.compile(
    r"\b("
    r"crea[rm]?\s+(?:un[ao]?\s+)?(?:nuevo\s+)?(?:excel|tabla|hoja|archivo|libro)|"
    r"hazme\s+(?:un[ao]?\s+)?(?:excel|tabla|hoja|archivo|plantilla)|"
    r"haz\s+(?:un[ao]?\s+)?(?:excel|tabla|hoja|archivo)|"
    r"genera[rm]?\s+(?:un[ao]?\s+)?(?:excel|tabla|hoja|archivo)|"
    r"necesito\s+(?:un[ao]?\s+)?(?:excel|tabla|hoja)\s+(?:con|para|de)|"
    r"quiero\s+(?:un[ao]?\s+)?(?:excel|tabla|hoja)\s+(?:con|para|de)"
    r")\b",
    re.IGNORECASE,
)

# Macros
_RE_GUARDAR_MACRO = re.compile(
    r"\b(?:guarda[r]?|crea[r]?|define[r]?)\s+(?:una?\s+)?macro\s+(?:llamada?\s+|con\s+nombre\s+)?['\"]?(\w+)['\"]?",
    re.IGNORECASE,
)
_RE_EJECUTAR_MACRO = re.compile(
    r"\b(?:aplica[r]?|ejecuta[r]?|usa[r]?|lanza[r]?)\s+(?:la\s+)?macro\s+['\"]?(\w+)['\"]?",
    re.IGNORECASE,
)
_RE_LISTAR_MACROS = re.compile(
    r"\b(?:lista[r]?\s+macros?|mis\s+macros?|qu[eé]\s+macros?\s+tengo|ver\s+macros?)\b",
    re.IGNORECASE,
)
_RE_BORRAR_MACRO = re.compile(
    r"\b(?:borra[r]?|elimina[r]?|quita[r]?)\s+(?:la\s+)?macro\s+['\"]?(\w+)['\"]?",
    re.IGNORECASE,
)

# Comparar archivos
_RE_COMPARAR = re.compile(
    r"\b("
    r"compara[r]?\s+(?:los?\s+)?(?:dos\s+)?archivos?|"
    r"diferencias?\s+entre\s+(?:los?\s+)?(?:dos\s+)?archivos?|"
    r"qu[eé]\s+(?:ha\s+)?cambiado|qu[eé]\s+cambios?\s+hay|"
    r"qu[eé]\s+diferencias?\s+(?:hay|tiene|existen)|"
    r"compara[r]?\s+(?:con\s+)?el\s+anterior|"
    r"diff\b|comparaci[oó]n\s+de\s+archivos?"
    r")\b",
    re.IGNORECASE,
)

# Previsualización de filas
_RE_PREVIEW = re.compile(
    r"\b(?:primeras?|[uú]ltimas?)\s+(\d+)\s*(?:filas?|registros?|l[ií]neas?|datos?)?|"
    r"\b(\d+)\s+(?:primeras?|[uú]ltimas?)\s*(?:filas?|registros?)?|"
    r"\bmu[eé]strame\s+(?:los?\s+)?datos\b|"
    r"\bprevisualiza[r]?\b|"
    r"\bver?\s+(?:los?\s+)?datos\b",
    re.IGNORECASE,
)

# Valores únicos
_RE_VALORES_UNICOS = re.compile(
    r"\b("
    r"valores?\s+[uú]nicos?|valores?\s+distintos?|"
    r"qu[eé]\s+(?:valores?|categor[ií]as?|opciones?|tipos?)\s+(?:hay|tiene|existen|aparecen)|"
    r"lista[r]?\s+(?:los?\s+)?(?:valores?|categor[ií]as?|opciones?)|"
    r"cu[aá]ntos?\s+(?:\w+\s+)?(?:distintos?|[uú]nicos?)\s+hay|"
    r"cu[aá]les?\s+son\s+los?\s+(?:distintos?|[uú]nicos?|posibles?)"
    r")\b",
    re.IGNORECASE,
)

# Explicar archivo
_RE_EXPLICAR_ARCHIVO = re.compile(
    r"\b("
    r"expl[ií]came\s+(?:este\s+)?(?:archivo|excel|datos?|tabla)|"
    r"qu[eé]\s+(?:contiene|hay en|tiene)\s+(?:este\s+)?(?:archivo|excel)|"
    r"descr[ií]beme\s+(?:este\s+)?(?:archivo|excel|datos?)|"
    r"resumen\s+(?:del?\s+)?archivo|analiza\s+(?:este\s+)?archivo|"
    r"de\s+qu[eé]\s+(?:va|trata)\s+(?:este\s+)?(?:archivo|excel)"
    r")\b",
    re.IGNORECASE,
)

# Exportar a CSV
_RE_EXPORTAR_CSV = re.compile(
    r"\b("
    r"exporta[r]?\s+(?:a\s+|como\s+|en\s+)?csv|"
    r"guarda[r]?\s+(?:como\s+|en\s+)?csv|"
    r"descarga[r]?\s+(?:en\s+|como\s+)?csv|"
    r"convierte[r]?\s+a\s+csv|"
    r"en\s+formato\s+csv|formato\s+csv"
    r")\b",
    re.IGNORECASE,
)

# Deshacer
_RE_UNDO = re.compile(
    r"\b("
    r"deshaz|deshacer|desh[aá]cer|"
    r"vuelve\s+atr[aá]s|volver\s+atr[aá]s|"
    r"revertir|revierte|rev[eé]rtelo|"
    r"undo|ctrl\s*\+?\s*z|"
    r"cancela\s+(?:el\s+)?(?:[uú]ltimo\s+)?cambio|"
    r"recupera\s+(?:el\s+)?(?:archivo|excel|datos?)\s+anterior"
    r")\b",
    re.IGNORECASE,
)

# Gráfico bajo demanda
_RE_GRAFICO = re.compile(
    r"\b("
    r"gr[aá]fico[s]?\b|chart\b|"
    r"dibuja[r]?\s+(?:un\s+)?gr[aá]fico|"
    r"pinta[r]?\s+(?:un\s+)?gr[aá]fico|"
    r"genera[r]?\s+(?:un\s+)?gr[aá]fico|"
    r"hazme\s+(?:un\s+)?gr[aá]fico|"
    r"muestra[r]?\s+(?:un\s+)?gr[aá]fico|"
    r"crea[r]?\s+(?:un\s+)?gr[aá]fico|"
    r"visualiza[r]?\s|representa[r]?\s+gr[aá]ficamente|"
    r"gr[aá]fico\s+de\s+(?:barras?|l[ií]neas?|sectores?|tarta|pie|dispersi[oó]n|scatter)|"
    r"histograma\b"
    r")\b",
    re.IGNORECASE,
)

# Análisis estadístico / correlaciones / tendencia
_RE_STATS = re.compile(
    r"\b("
    r"estad[ií]stica[s]?|distribuc[ií][oó]n|correlac[ií][oó]n[es]?|"
    r"an[áa]lisis\s+(?:completo|estad[ií]stico|de\s+datos)|"
    r"media\s+y\s+(?:mediana|desviaci[oó]n)|resumen\s+estad[ií]stico|"
    r"desviaci[oó]n\s+(?:est[aá]ndar|t[ií]pica)|percentil[es]?|"
    r"qu[eé]\s+columnas\s+(?:est[aá]n\s+)?(?:m[aá]s\s+)?relacionadas|"
    r"c[oó]mo\s+se\s+relacionan|mapa\s+de\s+calor|heatmap|"
    r"tendencia[s]?|proyecci[oó]n|evoluci[oó]n\s+de\s+|"
    r"c[oó]mo\s+(?:evolucionan|van\s+(?:mis\s+)?|crecen|bajan)|"
    r"predicci[oó]n|pron[oó]stico"
    r")\b",
    re.IGNORECASE,
)

# Tabla dinámica
_RE_TABLA_DINAMICA = re.compile(
    r"tabla[s]?\s+din[aá]mica[s]?",
    re.IGNORECASE,
)
_RE_SOLO_INFORMATIVA = re.compile(
    r"^[¿\s]*(qu[eé]\s+es\b|qu[eé]\s+son\b|c[oó]mo\s+(se\s+)?func|"
    r"para\s+qu[eé]\b|cu[aá]ndo\s+usar|por\s+qu[eé]\b|"
    r"explica[rm]?|qu[eé]\s+hace\b|qu[eé]\s+hacen\b|en\s+qu[eé]\s+consiste)",
    re.IGNORECASE,
)

# Sub-patrones para _analizar_estadisticas
_RE_CORR = re.compile(
    r"correlac[ií][oó]n|relacionad|mapa\s+de\s+calor|heatmap|c[oó]mo\s+se\s+relacionan",
    re.IGNORECASE,
)
_RE_TENDENCIA = re.compile(
    r"tendencia[s]?|proyecci[oó]n|evoluci[oó]n\s+de\s+|"
    r"c[oó]mo\s+(?:evolucionan|van\s+(?:mis\s+)?|crecen|bajan)|"
    r"predicci[oó]n|pron[oó]stico",
    re.IGNORECASE,
)

# Sentinel para respuestas de aclaración del LLM
_CLAVE_ACLARACION = "aclaracion_necesaria"
