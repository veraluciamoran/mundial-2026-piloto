# ============================================================
# 01_preparar_datos.R
# Piloto del Grupo A del Mundial 2026
#
# Objetivo:
# 1. Importar las tablas originales de Excel.
# 2. Validar identificadores.
# 3. Crear una fila por selección y partido.
# 4. Calcular resultados y puntos.
# 5. Calcular distancias entre sedes consecutivas.
# 6. Crear una tabla resumen por selección.
# 7. Exportar las tablas procesadas.
# ============================================================


# ------------------------------------------------------------
# 1. Cargar paquetes
# ------------------------------------------------------------

library(tidyverse)
library(readxl)
library(janitor)
library(lubridate)
library(geosphere)
library(here)


# ------------------------------------------------------------
# 2. Definir la ruta del Excel
# ------------------------------------------------------------

ruta_excel <- here(
  "data",
  "raw",
  "mundial_grupo_a.xlsx"
)

if (!file.exists(ruta_excel)) {
  stop(
    paste0(
      "No se encontró el archivo Excel en: ",
      ruta_excel,
      "\nComprueba que se llame mundial_grupo_a.xlsx ",
      "y esté dentro de data/raw/."
    )
  )
}


# ------------------------------------------------------------
# 3. Importar las hojas
# ------------------------------------------------------------

selecciones <- read_excel(
  path = ruta_excel,
  sheet = "selecciones"
) |>
  clean_names()

sedes <- read_excel(
  path = ruta_excel,
  sheet = "sedes"
) |>
  clean_names()

partidos <- read_excel(
  path = ruta_excel,
  sheet = "partidos"
) |>
  clean_names()

resultado_grupo <- read_excel(
  path = ruta_excel,
  sheet = "resultado_grupo"
) |>
  clean_names()

fuentes <- read_excel(
  path = ruta_excel,
  sheet = "fuentes"
) |>
  clean_names()


# ------------------------------------------------------------
# 4. Verificar nombres de columnas
# ------------------------------------------------------------

columnas_selecciones <- c(
  "seleccion_id",
  "seleccion",
  "grupo",
  "confederacion"
)

columnas_sedes <- c(
  "sede_id",
  "estadio",
  "ciudad",
  "pais",
  "latitud",
  "longitud"
)

columnas_partidos <- c(
  "partido_id",
  "fecha",
  "fase",
  "grupo",
  "seleccion_local_id",
  "seleccion_visitante_id",
  "sede_id",
  "goles_local",
  "goles_visitante"
)

columnas_resultado <- c(
  "seleccion_id",
  "posicion_grupo",
  "clasifico"
)

if (!all(columnas_selecciones %in% names(selecciones))) {
  stop(
    "La hoja selecciones no tiene todas las columnas requeridas."
  )
}

if (!all(columnas_sedes %in% names(sedes))) {
  stop(
    "La hoja sedes no tiene todas las columnas requeridas."
  )
}

if (!all(columnas_partidos %in% names(partidos))) {
  stop(
    "La hoja partidos no tiene todas las columnas requeridas."
  )
}

if (!all(columnas_resultado %in% names(resultado_grupo))) {
  stop(
    "La hoja resultado_grupo no tiene todas las columnas requeridas."
  )
}


# ------------------------------------------------------------
# 5. Verificar duplicados
# ------------------------------------------------------------

if (anyDuplicated(selecciones$seleccion_id) > 0) {
  stop(
    "Hay códigos duplicados en selecciones$seleccion_id."
  )
}

if (anyDuplicated(sedes$sede_id) > 0) {
  stop(
    "Hay códigos duplicados en sedes$sede_id."
  )
}

if (anyDuplicated(partidos$partido_id) > 0) {
  stop(
    "Hay códigos duplicados en partidos$partido_id."
  )
}

if (anyDuplicated(resultado_grupo$seleccion_id) > 0) {
  stop(
    "Hay códigos duplicados en resultado_grupo$seleccion_id."
  )
}


# ------------------------------------------------------------
# 6. Limpiar texto y tipos
# ------------------------------------------------------------

selecciones <- selecciones |>
  mutate(
    seleccion_id = str_trim(seleccion_id),
    seleccion = str_trim(seleccion),
    grupo = str_trim(grupo),
    confederacion = str_trim(confederacion)
  )

sedes <- sedes |>
  mutate(
    sede_id = str_trim(sede_id),
    estadio = str_trim(estadio),
    ciudad = str_trim(ciudad),
    pais = str_trim(pais),
    latitud = as.numeric(latitud),
    longitud = as.numeric(longitud)
  )

partidos <- partidos |>
  mutate(
    partido_id = str_trim(partido_id),
    fecha = as.Date(fecha),
    seleccion_local_id = str_trim(
      seleccion_local_id
    ),
    seleccion_visitante_id = str_trim(
      seleccion_visitante_id
    ),
    sede_id = str_trim(sede_id),
    goles_local = as.numeric(goles_local),
    goles_visitante = as.numeric(
      goles_visitante
    )
  )

resultado_grupo <- resultado_grupo |>
  mutate(
    seleccion_id = str_trim(seleccion_id),
    posicion_grupo = as.integer(
      posicion_grupo
    ),
    clasifico = str_trim(clasifico)
  )


# ------------------------------------------------------------
# 7. Validar coordenadas y fechas
# ------------------------------------------------------------

if (any(is.na(sedes$latitud))) {
  stop(
    "Hay valores no numéricos o vacíos en latitud."
  )
}

if (any(is.na(sedes$longitud))) {
  stop(
    "Hay valores no numéricos o vacíos en longitud."
  )
}

if (any(is.na(partidos$fecha))) {
  stop(
    "Hay fechas que R no pudo interpretar."
  )
}

if (any(is.na(partidos$goles_local)) ||
    any(is.na(partidos$goles_visitante))) {
  stop(
    "Hay resultados vacíos o no numéricos."
  )
}


# ------------------------------------------------------------
# 8. Validar códigos de selecciones en partidos
# ------------------------------------------------------------

codigos_validos <- selecciones$seleccion_id

codigos_partidos <- unique(
  c(
    partidos$seleccion_local_id,
    partidos$seleccion_visitante_id
  )
)

codigos_inexistentes <- setdiff(
  codigos_partidos,
  codigos_validos
)

if (length(codigos_inexistentes) > 0) {
  stop(
    paste0(
      "Estos códigos aparecen en partidos pero no ",
      "en selecciones: ",
      paste(
        codigos_inexistentes,
        collapse = ", "
      )
    )
  )
}


# ------------------------------------------------------------
# 9. Validar códigos de sedes en partidos
# ------------------------------------------------------------

sedes_inexistentes <- setdiff(
  partidos$sede_id,
  sedes$sede_id
)

if (length(sedes_inexistentes) > 0) {
  stop(
    paste0(
      "Estas sedes aparecen en partidos pero no ",
      "en sedes: ",
      paste(
        sedes_inexistentes,
        collapse = ", "
      )
    )
  )
}


# ------------------------------------------------------------
# 10. Crear perspectiva del equipo local
# ------------------------------------------------------------

partidos_local <- partidos |>
  transmute(
    partido_id,
    fecha,
    fase,
    grupo,
    seleccion_id = seleccion_local_id,
    oponente_id = seleccion_visitante_id,
    sede_id,
    condicion = "Local",
    goles_favor = goles_local,
    goles_contra = goles_visitante
  )


# ------------------------------------------------------------
# 11. Crear perspectiva del equipo visitante
# ------------------------------------------------------------

partidos_visitante <- partidos |>
  transmute(
    partido_id,
    fecha,
    fase,
    grupo,
    seleccion_id = seleccion_visitante_id,
    oponente_id = seleccion_local_id,
    sede_id,
    condicion = "Visitante",
    goles_favor = goles_visitante,
    goles_contra = goles_local
  )


# ------------------------------------------------------------
# 12. Unir ambas perspectivas
# ------------------------------------------------------------

partidos_largos <- bind_rows(
  partidos_local,
  partidos_visitante
) |>
  mutate(
    resultado = case_when(
      goles_favor > goles_contra ~ "Victoria",
      goles_favor == goles_contra ~ "Empate",
      goles_favor < goles_contra ~ "Derrota"
    ),
    puntos = case_when(
      resultado == "Victoria" ~ 3,
      resultado == "Empate" ~ 1,
      resultado == "Derrota" ~ 0
    ),
    diferencia_goles = (
      goles_favor - goles_contra
    )
  ) |>
  left_join(
    selecciones |>
      select(
        seleccion_id,
        seleccion,
        confederacion
      ),
    by = "seleccion_id"
  ) |>
  left_join(
    selecciones |>
      select(
        oponente_id = seleccion_id,
        oponente = seleccion
      ),
    by = "oponente_id"
  ) |>
  left_join(
    sedes,
    by = "sede_id"
  ) |>
  arrange(
    seleccion_id,
    fecha,
    partido_id
  )


# ------------------------------------------------------------
# 13. Verificar las uniones
# ------------------------------------------------------------

if (any(is.na(partidos_largos$seleccion))) {
  stop(
    "Al menos una selección no pudo unirse correctamente."
  )
}

if (any(is.na(partidos_largos$oponente))) {
  stop(
    "Al menos un oponente no pudo unirse correctamente."
  )
}

if (any(is.na(partidos_largos$ciudad))) {
  stop(
    "Al menos una sede no pudo unirse correctamente."
  )
}


# ------------------------------------------------------------
# 14. Calcular desplazamientos entre sedes
# ------------------------------------------------------------

recorridos <- partidos_largos |>
  group_by(seleccion_id) |>
  arrange(
    fecha,
    partido_id,
    .by_group = TRUE
  ) |>
  mutate(
    numero_partido = row_number(),
    
    latitud_anterior = lag(latitud),
    
    longitud_anterior = lag(longitud),
    
    ciudad_anterior = lag(ciudad),
    
    distancia_tramo_km = case_when(
      numero_partido == 1 ~ 0,
      TRUE ~ geosphere::distHaversine(
        p1 = cbind(
          longitud_anterior,
          latitud_anterior
        ),
        p2 = cbind(
          longitud,
          latitud
        )
      ) / 1000
    ),
    
    dias_descanso = as.numeric(
      fecha - lag(fecha)
    )
  ) |>
  ungroup() |>
  mutate(
    distancia_tramo_km = round(
      distancia_tramo_km,
      digits = 1
    )
  )


# ------------------------------------------------------------
# 15. Crear resumen calculado por selección
# ------------------------------------------------------------

resumen_calculado <- recorridos |>
  group_by(
    seleccion_id,
    seleccion,
    grupo,
    confederacion
  ) |>
  summarise(
    partidos = n(),
    
    puntos = sum(
      puntos,
      na.rm = TRUE
    ),
    
    goles_favor = sum(
      goles_favor,
      na.rm = TRUE
    ),
    
    goles_contra = sum(
      goles_contra,
      na.rm = TRUE
    ),
    
    diferencia_goles = sum(
      diferencia_goles,
      na.rm = TRUE
    ),
    
    distancia_entre_sedes_km = sum(
      distancia_tramo_km,
      na.rm = TRUE
    ),
    
    mayor_tramo_km = max(
      distancia_tramo_km,
      na.rm = TRUE
    ),
    
    cantidad_cambios_sede = sum(
      distancia_tramo_km > 0,
      na.rm = TRUE
    ),
    
    descanso_promedio_dias = mean(
      dias_descanso,
      na.rm = TRUE
    ),
    
    .groups = "drop"
  )


# ------------------------------------------------------------
# 16. Incorporar resultado oficial del grupo
# ------------------------------------------------------------

resumen_selecciones <- resumen_calculado |>
  left_join(
    resultado_grupo,
    by = "seleccion_id"
  ) |>
  arrange(posicion_grupo)


# ------------------------------------------------------------
# 17. Verificaciones deportivas
# ------------------------------------------------------------

if (!all(
  resumen_selecciones$partidos == 3
)) {
  warning(
    "Al menos una selección no tiene exactamente tres partidos."
  )
}

if (sum(resumen_selecciones$puntos) != 17) {
  warning(
    paste0(
      "La suma total de puntos no es 17. ",
      "Comprueba los resultados."
    )
  )
}


# ------------------------------------------------------------
# 18. Crear carpeta processed si no existe
# ------------------------------------------------------------

dir.create(
  here("data", "processed"),
  recursive = TRUE,
  showWarnings = FALSE
)


# ------------------------------------------------------------
# 19. Exportar tablas
# ------------------------------------------------------------

write_csv(
  partidos_largos,
  here(
    "data",
    "processed",
    "partidos_largos.csv"
  )
)

write_csv(
  recorridos,
  here(
    "data",
    "processed",
    "recorridos.csv"
  )
)

write_csv(
  resumen_selecciones,
  here(
    "data",
    "processed",
    "resumen_selecciones.csv"
  )
)


# ------------------------------------------------------------
# 20. Mensaje final
# ------------------------------------------------------------

message(
  "Proceso completado correctamente."
)

print(
  resumen_selecciones |>
    select(
      posicion_grupo,
      seleccion,
      puntos,
      goles_favor,
      goles_contra,
      diferencia_goles,
      distancia_entre_sedes_km,
      clasifico
    )
)