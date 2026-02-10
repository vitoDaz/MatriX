# 1. CARGA DE LIBRERÍAS (Sin duplicados)
library(tidyverse) 
library(readxl)
library(writexl)
library(janitor)
library(tidyxl)
library(skimr)

# 2. CONFIGURACIÓN DE ENTORNO
setwd(r"(/home/vitodaz/Documentos/INGENIERIA/2026-02-08/DB/Avance II/MatrixXx3)")


# 3. LECTURA DE DATOS PRINCIPALES
# Solo trabajaremos con estos dos archivos
matrixtec <- read_excel("Matriz técnica general Andes Norte (CC-130).xlsx", skip = 20)
Avancedf  <- read_excel("Avance CC-130.xlsx", skip = 20)

# 4. EXTRACCIÓN DE METADATOS DE COLOR (Optimizado)
archivo_avance <- "Avance CC-130.xlsx"
df_completo    <- xlsx_cells(archivo_avance)
formatos       <- xlsx_formats(archivo_avance)

# Identificar celdas coloreadas desde columna U (21) en adelante
celdas_coloreadas <- df_completo %>%
  filter(
    col >= 21,    # Columna U en adelante
    row >= 22     # Fila 22 en adelante (asumiendo que la 21 es el encabezado)
  ) %>% 
  mutate(
    color_relleno = formatos$local$fill$patternFill$fgColor$rgb[local_format_id]
  ) %>%
  filter(!is.na(color_relleno))

#  ESTANDARIZACIÓN Y LIMPIEZA (Diccionario Optimizado)
# Diccionario para renombrar columnas a nombres cortos
diccionario_final <- c(
  "bulto_interno"     = "n_bulto_interno",
  "contrato"          = "contrato",
  "prov"              = "prov",
  "tipo_suministro"   = "tipo_de_suministro_suministro_principal_suministro_despiece_repuesto_capitalizable_repuesto_pem_repuesto_1_o_2_anos_de_op",
  "txt_documento"     = "txt_de_documento",
  "txt_tecnico"       = "txt_tecnico",
  "mm"                = "mm",
  "pwp"               = "pwp",
  "num_serie_parte"   = "numero_serie_parte",
  "tag_elemento"      = "tag_del_elemento",
  "tag"               = "tag",
  "bulto_proveedor"   = "n_bulto_proveedor",
  "cantidad"          = "cantidad",
  "sobrante_faltante" = "sobrante_faltante",
  "unidad_medida"     = "unidad_de_medida",
  "lugar"             = "lugar",
  "acopio"            = "acopio",
  "tipo_embalaje"     = "tipo_de_embalaje",
  "area"              = "area",
  "sub_area"          = "sub_area",
  "nivel_wbs"         = "nivel_wbs",
  "sistema"           = "sistema",
  "conjunto"          = "conjunto",
  "elemento"          = "elemento",
  "instrumentos"      = "instrumentos",
  "pwp2"              = "pwp2",
  "especialidad"      = "especialidad_elec_mec_emec"
)

matrixtec_clean <- matrixtec %>% 
  clean_names() %>% 
  rename(any_of(diccionario_final)) %>%
  # NUEVO: Forzar cantidad a número y limpiar espacios
  mutate(cantidad = as.numeric(cantidad))

Avancedf_clean <- Avancedf %>% 
  clean_names() %>% 
  rename(any_of(diccionario_final)) %>%
  # NUEVO: Forzar cantidad a número y limpiar espacios
  mutate(cantidad = as.numeric(cantidad))

# Paso A: Crear un "Diccionario Maestro de Lugares" desde la Matriz Técnica
# Agrupamos por bulto y tomamos el primer lugar válido que encontremos.
referencia_lugar <- matrixtec_clean %>%
  group_by(bulto_interno) %>%  # <--- CORREGIDO: Usamos el nombre ya limpio
  summarise(
    lugar_ref = first(na.omit(lugar)), # Toma el primer lugar no nulo encontrado
    .groups = "drop"
  ) %>%
  filter(!is.na(bulto_interno)) # Eliminamos filas vacías si las hubiera

# Cruzamos usando la columna común 'bulto_interno'
Avancedf_final <- Avancedf_clean %>%
  left_join(referencia_lugar, by = "bulto_interno") %>%
  mutate(
    # Si Avance ya tiene lugar, lo mantiene. Si no, usa el de la referencia.
    lugar = coalesce(lugar, lugar_ref) 
  ) %>%
  select(-lugar_ref) # Eliminamos la columna auxiliar

# Verificación rápida
print(paste("Filas cruzadas:", nrow(Avancedf_final)))
print(glimpse(Avancedf_final))
##########################################################################################################
# 7. HOMOGENEIZACIÓN DE ESTRUCTURAS
# Identificar qué columnas tiene el objetivo (Avance)
columnas_objetivo <- colnames(Avancedf_final)

# Creamos una versión de Matrix que solo tiene las columnas que nos interesan
matrixtec_final <- matrixtec_clean %>%
  select(any_of(columnas_objetivo))

#  Obtenemos una fila vacía con los tipos de datos CORRECTOS (de Avance)
template_tipos <- Avancedf_final %>% filter(FALSE)

#  Forzamos a que las columnas de matrixtec_final tengan el mismo tipo que el template
# Esto convierte, por ejemplo, character a double o viceversa según sea necesario
for (col_name in intersect(names(matrixtec_final), names(template_tipos))) {
  class(matrixtec_final[[col_name]]) <- class(template_tipos[[col_name]])
}

#  Ahora sí, unimos las columnas faltantes (que vendrán de Avancedf_final)
matrixtec_final <- matrixtec_final %>%
  bind_rows(template_tipos) %>%
  # Reordenamos las columnas para que el espejo sea perfecto
  select(all_of(columnas_objetivo))


glimpse(matrixtec_final)
glimpse(Avancedf_final)

skim(Avancedf_final)
skim(matrixtec_final)


write_xlsx(Avancedf_final, "Avancedf_final.xlsx")
write_xlsx(matrixtec_final, "matrixtec_final.xlsx")
