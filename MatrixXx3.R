# 1. CARGA DE LIBRERÍAS
library(tidyverse) 
library(readxl)
library(writexl)
library(janitor)
library(tidyxl)
library(skimr)
library(httr) 


# 2. CONFIGURACIÓN DE RUTAS DINÁMICAS (Windows/Linux)
# Detecta si estás en el ThinkPad (Linux) o en Windows
if (Sys.info()["sysname"] == "Linux") {
  base_path <- "/home/vitodaz/Documentos/INGENIERIA/2026-02-08/DB/Avance II/MatrixXx3"
} else {
  # Ajusta esta ruta a tu carpeta de Windows
  base_path <- "C:/Users/vitodaz/Documents/MatrixXx3" 
}

if (!dir.exists(base_path)) dir.create(base_path, recursive = TRUE)
setwd(base_path)

# 3. DESCARGA DE ARCHIVOS DESDE DROPBOX
# Cambiamos dl=0 por dl=1 para forzar la descarga directa
url_avance <- "https://www.dropbox.com/scl/fi/z3j8t9y1lkv01ipfsvv55/Avance-CC-130.xlsx?rlkey=lif1tl03m4yi1kxh5zzxd6i15&st=7obi660x&dl=1"
url_matrix <- "https://www.dropbox.com/scl/fi/8jgu635nt0obntp7ee3rs/Matriz-t-cnica-general-Andes-Norte-CC-130.xlsx?rlkey=ce4r0xp5blkzvkqi7qk07aaqk&st=6kk1i3r3&dl=1"

# Descargamos los archivos temporalmente para que read_excel y tidyxl funcionen
archivo_avance <- "Avance_CC_130.xlsx"
archivo_matrix <- "Matriz_tecnica_Andes_Norte.xlsx"

GET(url_avance, write_disk(archivo_avance, overwrite = TRUE))
GET(url_matrix, write_disk(archivo_matrix, overwrite = TRUE))

# 4. LECTURA DE DATOS
matrixtec <- read_excel(archivo_matrix, skip = 20)
Avancedf  <- read_excel(archivo_avance, skip = 20)

# 5. EXTRACCIÓN DE METADATOS DE COLOR
df_completo <- xlsx_cells(archivo_avance)
formatos    <- xlsx_formats(archivo_avance)

celdas_coloreadas <- df_completo %>%
  filter(
    col >= 21,    # Columna U en adelante
    row >= 22     # Fila 22 en adelante
  ) %>% 
  mutate(
    color_relleno = formatos$local$fill$patternFill$fgColor$rgb[local_format_id]
  ) %>%
  filter(!is.na(color_relleno))

# 6. ESTANDARIZACIÓN Y LIMPIEZA
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
  mutate(cantidad = as.numeric(cantidad))

Avancedf_clean <- Avancedf %>% 
  clean_names() %>% 
  rename(any_of(diccionario_final)) %>%
  mutate(cantidad = as.numeric(cantidad))

# 7. CRUCE DE LUGARES (Diccionario Maestro)
referencia_lugar <- matrixtec_clean %>%
  group_by(bulto_interno) %>% 
  summarise(
    lugar_ref = first(na.omit(lugar)),
    .groups = "drop"
  ) %>%
  filter(!is.na(bulto_interno))

Avancedf_final <- Avancedf_clean %>%
  left_join(referencia_lugar, by = "bulto_interno") %>%
  mutate(lugar = coalesce(lugar, lugar_ref)) %>%
  select(-lugar_ref)

# 8. HOMOGENEIZACIÓN DE ESTRUCTURAS
columnas_objetivo <- colnames(Avancedf_final)

matrixtec_final <- matrixtec_clean %>%
  select(any_of(columnas_objetivo))

# Sincronización de tipos de datos
for (col_name in intersect(names(matrixtec_final), names(Avancedf_final))) {
  target_class <- class(Avancedf_final[[col_name]])
  matrixtec_final[[col_name]] <- as(matrixtec_final[[col_name]], target_class)
}
View(Avancedf_final)
View(matrixtec_final)
# 9. EXPORTACIÓN FINAL
write_xlsx(Avancedf_final, "Avancedf_final.xlsx")
write_xlsx(matrixtec_final, "matrixtec_final.xlsx")

print("Proceso completado con éxito.")
