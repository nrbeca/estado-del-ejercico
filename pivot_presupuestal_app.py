"""
Constructor de Tabla Dinámica Presupuestal — MAP / SICOP
==========================================================
App Streamlit que integra los reportes crudos de MAP y SICOP, deja armar
cualquier reporte tipo tabla dinámica (agregar/quitar filas, columnas y
valores, filtrar por cualquier campo de la base — incluida Unidad
Responsable, Partida y Programa por nombre) y descarga el resultado en
Excel con el formato institucional del "Estado del Ejercicio".

Cómo correrla:
    pip install -r requirements.txt
    streamlit run pivot_presupuestal_app.py

Qué hace por ti automáticamente al cargar un archivo crudo:
    - Detecta la codificación (MAP = utf-8, SICOP = latin-1) sola.
    - Construye la Partida completa en SICOP (Capítulo+Concepto+Genérica+
      Específica) igual que en tus reportes actuales.
    - Junta los catálogos de catalogs/unidades.csv, catalogs/partidas.csv
      y catalogs/programas.csv para mostrar nombres, no solo códigos, en
      los filtros y en las filas del reporte. Si un código no está en el
      catálogo, muestra el código tal cual — puedes ir agregando filas a
      esos CSV para completar la cobertura.
    - Calcula, para cada familia de importes (Original, Modificado,
      Comprometido, Ejercido, Reservas, etc.), el total "Anual" y el
      "Al periodo" (acumulado de enero al mes que elijas), igual que en
      el Estado del Ejercicio.
    - Calcula el Importe Disponible como Modificado − Ejercido − Comprometido
      (fórmula verificada contra tu archivo de ejemplo, cuadra al centavo).
"""

from __future__ import annotations

import io
import re
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Paleta y estilos — tomados directamente de formato_estado_del_ejercicio.xlsx
# ---------------------------------------------------------------------------
BURDEOS = "621333"     # encabezados de columnas
VERDE = "1E5B4F"       # encabezados de grupo ("Anual" / "Al periodo")
GRIS_TOTAL = "D9D9D9"  # fila de Total general
FORMATO_MONEDA = '_-"$"* #,##0.00_-;\\-"$"* #,##0.00_-;_-"$"* "-"??_-;_-@_-'

THIN_GRIS = Side(style="thin", color="808080")
BORDE = Border(left=THIN_GRIS, right=THIN_GRIS, top=THIN_GRIS, bottom=THIN_GRIS)
THIN_BLANCO = Side(style="thin", color="FFFFFF")
BORDE_ENCABEZADO = Border(left=THIN_BLANCO, right=THIN_BLANCO, top=THIN_BLANCO, bottom=THIN_BLANCO)

MESES = ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"]
NOMBRES_MES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
               "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

# ---------------------------------------------------------------------------
# Catálogos embebidos (código -> nombre), tomados de config.py (nrbeca/nuevo).
# Van directamente en el código para no depender de una carpeta catalogs/
# aparte (que además el .gitignore del repo excluye por ser *.csv).
# ---------------------------------------------------------------------------
CATALOGO_UNIDADES = {
    '100': 'Secretaría',
    '106': 'Coordinación de Legislación y Consulta',
    '107': 'Coordinación de lo Contencioso',
    '110': 'Unidad de Asuntos Jurídicos',
    '111': 'Dirección General de Comunicación Social',
    '112': 'Coordinación de Atención Legislativa',
    '117': 'Coordinación de Asuntos Internacionales',
    '119': 'Dirección General de Planeación y Evaluación de Políticas y Programas',
    '120': 'Dirección General del Servicio de Información Agroalimentaria y Pesquera',
    '200': 'Subsecretaría de Agricultura y Desarrollo Rural',
    '210': 'Dirección General de Logística y Alimentación',
    '211': 'Dirección General de Normalización Agroalimentaria',
    '212': 'Dirección General de Organización para la Productividad',
    '213': 'Coordinación General de Promoción Comercial y Fomento a las Exportaciones',
    '214': 'Dirección General de la Autosuficiencia Alimentaria',
    '215': 'Coordinación General de Enlace Sectorial',
    '220': 'Unidad de Bienestar para el Campo',
    '221': 'Dirección General de Fertilizantes para el Bienestar',
    '222': 'Dirección General de Producción para el Bienestar',
    '223': 'Dirección General de Ordenamiento Pesquero y Acuícola',
    '224': 'Dirección General de Inspección y Vigilancia',
    '225': 'Coordinación General de Delegaciones',
    '226': 'Órgano Interno de Control',
    '227': 'Coordinación General de Asuntos Internacionales',
    '230': 'Dirección General de Administración y Finanzas',
    '231': 'Coordinación General de Agricultura',
    '232': 'Delegación de Programas para el Desarrollo',
    '233': 'Dirección General de Agregación de Valor y Comercialización',
    '235': 'Dirección General de Programación, Presupuesto y Finanzas',
    '237': 'Dirección General de Tecnologías de la Información y Comunicaciones',
    '240': 'Coordinación General de Innovación y Transición Agroecológica',
    '242': 'Dirección General de Transición Agroecológica',
    '245': 'Dirección General de Recursos Materiales, Inmuebles y Servicios',
    '250': 'Unidad de Operación Territorial y Eficiencia Hídrica Agroalimentaria',
    '252': 'Dirección General de Intervención Territorial Estratégica',
    '253': 'Dirección General de Eficacia Hídrica en Riego y Temporal',
    '260': 'Oficina de Representación en Aguascalientes',
    '261': 'Oficina de Representación en Baja California',
    '262': 'Oficina de Representación en Baja California Sur',
    '263': 'Oficina de Representación en Campeche',
    '264': 'Oficina de Representación en Coahuila',
    '265': 'Oficina de Representación en Colima',
    '266': 'Oficina de Representación en Chiapas',
    '267': 'Oficina de Representación en Chihuahua',
    '268': 'Oficina de Representación en la Ciudad de México',
    '269': 'Oficina de Representación en Durango',
    '270': 'Oficina de Representación en Guanajuato',
    '271': 'Oficina de Representación en Guerrero',
    '272': 'Oficina de Representación en Hidalgo',
    '273': 'Oficina de Representación en Jalisco',
    '274': 'Oficina de Representación en el Estado de México',
    '275': 'Oficina de Representación en Michoacán',
    '276': 'Oficina de Representación en Morelos',
    '277': 'Oficina de Representación en Nayarit',
    '278': 'Oficina de Representación en Nuevo León',
    '279': 'Oficina de Representación en Oaxaca',
    '280': 'Oficina de Representación en Puebla',
    '281': 'Oficina de Representación en Querétaro',
    '282': 'Oficina de Representación en Quintana Roo',
    '283': 'Oficina de Representación en San Luis Potosí',
    '284': 'Oficina de Representación en Sinaloa',
    '285': 'Oficina de Representación en Sonora',
    '286': 'Oficina de Representación en Tabasco',
    '287': 'Oficina de Representación en Tamaulipas',
    '288': 'Oficina de Representación en Tlaxcala',
    '289': 'Oficina de Representación en Veracruz',
    '290': 'Oficina de Representación en Yucatán',
    '291': 'Oficina de Representación en Zacatecas',
    '292': 'Oficina de Representación en la Región Lagunera',
    '410': 'Dirección General de Fortalecimiento a la Agricultura Familiar',
    '411': 'Dirección General de Integración Económica',
    '413': 'Dirección General de Investigación, Desarrollo Tecnológico y Extensionismo',
    '500': 'Unidad de Administración y Finanzas',
    '510': 'Dirección General de Programación, Presupuesto y Finanzas',
    '511': 'Dirección General de Capital Humano y Desarrollo Organizacional',
    '512': 'Dirección General de Recursos Materiales, Inmuebles y Servicios',
    '513': 'Dirección General de Tecnologías de la Información y Comunicaciones',
    '810': 'SENASICA',
    '811': 'SNICS',
    '900': 'Coordinación General de Producción, Comercialización, Sustentabilidad e Innovación',
    '910': 'Unidad de Innovación, Sustentabilidad y Resiliencia Climática',
    '911': 'Dirección General de Desarrollo e Innovación',
    '912': 'Dirección General de Sustentabilidad y Resiliencia Climática',
    '920': 'Unidad de Producción, Comercialización y Financiamiento',
    '921': 'Dirección General de Producción Agrícola',
    '922': 'Dirección General de Producción Ganadera, Pesquera y Acuícola',
    '923': 'Dirección General de Precios, Ordenamiento Comercial y Valor Agregado',
    '924': 'Dirección General de Financiamiento y Gestión de Riesgos',
    'A1I': 'Universidad Autónoma Chapingo',
    'AFU': 'Comité Nacional para el Desarrollo Sustentable de la Caña de Azúcar',
    'B00': 'Servicio Nacional de Sanidad, Inocuidad y Calidad Agroalimentaria',
    'C00': 'Servicio Nacional de Inspección y Certificación de Semillas',
    'D00': 'Colegio Superior Agropecuario del Estado de Guerrero',
    'I00': 'Comisión Nacional de Acuacultura y Pesca',
    'I6L': 'Fideicomiso de Riesgo Compartido',
    'I9H': 'Instituto Nacional para el Desarrollo de Capacidades del Sector Rural, A.C.',
    'IZC': 'Colegio de Postgraduados',
    'IZI': 'Comisión Nacional de las Zonas Áridas',
    'JAG': 'Instituto Nacional de Investigaciones Forestales, Agrícolas y Pecuarias',
    'JAL': 'Productora de Semillas para el Bienestar',
    'JBK': 'Productora Nacional de Biológicos Veterinarios',
    'RJL': 'Instituto Mexicano de Investigación en Pesca y Acuacultura Sustentables',
    'VSS': 'Alimentación para el Bienestar, S.A de C.V.',
    'VST': 'Leche para el Bienestar, S.A. de C.V.',
}

CATALOGO_PARTIDAS = {
    '11101': 'Dietas (Ramos Autónomos)',
    '11201': 'Haberes',
    '11301': 'Sueldos base',
    '11401': 'Retribuciones por adscripción en el extranjero',
    '12101': 'Honorarios',
    '12201': 'Remuneraciones al personal eventual',
    '12202': 'Compensaciones a sustitutos de profesores',
    '12301': 'Retribuciones por servicios en período de formación profesional',
    '12401': 'Retribución a los representantes de los trabajadores y de los patrones en la Junta Federal de Conciliación y Arbitraje',
    '13101': 'Prima quinquenal por años de servicios efectivos prestados',
    '13102': 'Acreditación por años de servicio en la docencia y al personal administrativo de las instituciones de educación superior',
    '13103': 'Prima de perseverancia por años de servicio activo en el Ejército, Fuerza Aérea y Armada Mexicanos',
    '13104': 'Antigüedad',
    '13201': 'Primas de vacaciones y dominical',
    '13202': 'Aguinaldo o gratificación de fin de año',
    '13204': 'Primas de vacaciones y dominical de áreas administrativas (Ramos Autónomos)',
    '13301': 'Remuneraciones por horas extraordinarias',
    '13401': 'Acreditación por titulación en la docencia',
    '13402': 'Acreditación al personal docente por años de estudio de licenciatura',
    '13403': 'Compensaciones por servicios especiales',
    '13404': 'Compensaciones por servicios eventuales',
    '13405': 'Compensaciones de retiro',
    '13406': 'Compensaciones de servicios',
    '13407': 'Compensaciones adicionales por servicios especiales',
    '13408': 'Asignaciones docentes, pedagógicas genéricas y específicas',
    '13409': 'Compensación por adquisición de material didáctico',
    '13410': 'Compensación por actualización y formación académica',
    '13411': 'Compensaciones a médicos residentes',
    '13412': 'Gastos contingentes para el personal radicado en el extranjero',
    '13413': 'Asignaciones para la conclusión de servicios en la Administración Pública Federal',
    '13414': 'Asignaciones conforme al régimen laboral',
    '13501': 'Sobrehaberes',
    '13601': 'Asignaciones de técnico',
    '13602': 'Asignaciones de mando',
    '13603': 'Asignaciones por comisión',
    '13604': 'Asignaciones de vuelo',
    '13605': 'Asignaciones de técnico especial',
    '13701': 'Honorarios especiales',
    '13801': 'Participaciones por vigilancia en el cumplimiento de las leyes y custodia de valores',
    '14101': 'Aportaciones al ISSSTE',
    '14102': 'Aportaciones al ISSFAM',
    '14103': 'Aportaciones al IMSS',
    '14104': 'Aportaciones de seguridad social contractuales',
    '14105': 'Aportaciones al seguro de cesantía en edad avanzada y vejez',
    '14201': 'Aportaciones al FOVISSSTE',
    '14202': 'Aportaciones al INFONAVIT',
    '14301': 'Aportaciones al Sistema de Ahorro para el Retiro',
    '14302': 'Depósitos para el ahorro solidario',
    '14401': 'Cuotas para el seguro de vida del personal civil',
    '14402': 'Cuotas para el seguro de vida del personal militar',
    '14403': 'Cuotas para el seguro de gastos médicos del personal civil',
    '14404': 'Cuotas para el seguro de separación individualizado',
    '14405': 'Cuotas para el seguro colectivo de retiro',
    '14406': 'Seguro de responsabilidad civil, asistencia legal y otros seguros',
    '15101': 'Cuotas para el fondo de ahorro del personal civil',
    '15102': 'Cuotas para el fondo de ahorro de generales, almirantes, jefes y oficiales',
    '15103': 'Cuotas para el fondo de trabajo del personal del Ejército, Fuerza Aérea y Armada Mexicanos',
    '15201': 'Indemnizaciones por accidentes en el trabajo',
    '15202': 'Pago de liquidaciones',
    '15203': 'Fondo para indemnizaciones (Ramos Autónomos)',
    '15301': 'Prestaciones de retiro',
    '15302': 'Prestaciones y previsiones de retiro (Ramos Autónomos)',
    '15401': 'Prestaciones establecidas por condiciones generales de trabajo o contratos colectivos de trabajo',
    '15402': 'Compensación garantizada',
    '15403': 'Asignaciones adicionales al sueldo',
    '15405': 'Compensación de Apoyo (Ramos Autónomos)',
    '15501': 'Apoyos a la capacitación de los servidores públicos',
    '15901': 'Otras prestaciones',
    '15902': 'Pago extraordinario por riesgo',
    '16101': 'Incrementos a las percepciones',
    '16102': 'Creación de plazas',
    '16103': 'Otras medidas de carácter laboral y económico',
    '16104': 'Previsiones para aportaciones al ISSSTE',
    '16105': 'Previsiones para aportaciones al FOVISSSTE',
    '16106': 'Previsiones para aportaciones al Sistema de Ahorro para el Retiro',
    '16107': 'Previsiones para aportaciones al seguro de cesantía en edad avanzada y vejez',
    '16108': 'Previsiones para los depósitos al ahorro solidario',
    '16109': 'Previsiones por adecuaciones a las estructuras ocupacionales',
    '17101': 'Estímulos por productividad y eficiencia',
    '17102': 'Estímulos al personal operativo',
    '21101': 'Materiales y útiles de oficina',
    '21102': 'Material electoral (Ramos Autónomos)',
    '21199': 'Materiales de administración, emisión de documentos y artículos oficiales',
    '21201': 'Materiales y útiles de impresión y reproducción',
    '21301': 'Material estadístico y geográfico',
    '21401': 'Materiales y útiles consumibles para el procesamiento en equipos y bienes informáticos',
    '21501': 'Material de apoyo informativo',
    '21502': 'Material para información en actividades de investigación científica y tecnológica',
    '21601': 'Material de limpieza',
    '21701': 'Materiales y suministros para planteles educativos',
    '21801': 'Materiales para el registro e identificación de bienes y personas',
    '22101': 'Productos alimenticios para el Ejército, Fuerza Aérea y Armada Mexicanos',
    '22102': 'Productos alimenticios para personas derivado de la prestación de servicios públicos',
    '22103': 'Productos alimenticios para el personal que realiza labores en campo',
    '22104': 'Productos alimenticios para el personal en las instalaciones',
    '22105': 'Productos alimenticios para la población en caso de desastres naturales',
    '22106': 'Productos alimenticios para el personal derivado de actividades extraordinarias',
    '22199': 'Alimentos y utensilios',
    '22201': 'Productos alimenticios para animales',
    '22301': 'Utensilios para el servicio de alimentación',
    '23101': 'Productos alimenticios, agropecuarios y forestales adquiridos como materia prima',
    '23199': 'Materias primas y materiales de producción y comercialización',
    '23201': 'Insumos textiles adquiridos como materia prima',
    '23301': 'Productos de papel, cartón e impresos adquiridos como materia prima',
    '23401': 'Combustibles, lubricantes, aditivos, carbón y sus derivados adquiridos como materia prima',
    '23501': 'Productos químicos, farmacéuticos y de laboratorio adquiridos como materia prima',
    '23601': 'Productos metálicos y a base de minerales no metálicos adquiridos como materia prima',
    '23701': 'Productos de cuero, piel, plástico y hule adquiridos como materia prima',
    '23801': 'Mercancías para su comercialización en tiendas del sector público',
    '23901': 'Otros productos adquiridos como materia prima',
    '23902': 'Petróleo, gas y sus derivados adquiridos como materia prima',
    '24101': 'Productos minerales no metálicos',
    '24199': 'Materiales y artículos de construcción y de reparación',
    '24201': 'Cemento y productos de concreto',
    '24301': 'Cal, yeso y productos de yeso',
    '24401': 'Madera y productos de madera',
    '24501': 'Vidrio y productos de vidrio',
    '24601': 'Material eléctrico y electrónico',
    '24701': 'Artículos metálicos para la construcción',
    '24801': 'Materiales complementarios',
    '24901': 'Otros materiales y artículos de construcción y reparación',
    '25101': 'Productos químicos básicos',
    '25199': 'Productos químicos, farmacéuticos y de laboratorio',
    '25201': 'Plaguicidas, abonos y fertilizantes',
    '25301': 'Medicinas y productos farmacéuticos',
    '25401': 'Materiales, accesorios y suministros médicos',
    '25501': 'Materiales, accesorios y suministros de laboratorio',
    '25601': 'Fibras sintéticas, hules, plásticos y derivados',
    '25901': 'Otros productos químicos',
    '26101': 'Combustibles para programas de seguridad pública y nacional',
    '26102': 'Combustibles para servicios públicos y operación de programas',
    '26103': 'Combustibles para servicios administrativos',
    '26104': 'Combustibles asignados a servidores públicos',
    '26105': 'Combustibles para maquinaria y equipo de producción',
    '26106': 'PIDIREGAS cargos variables',
    '26107': 'Combustibles nacionales para plantas productivas',
    '26108': 'Combustibles de importación para plantas productivas',
    '26199': 'Combustibles, lubricantes y aditivos',
    '27101': 'Vestuario y uniformes',
    '27199': 'Vestuario, blancos, prendas de protección y artículos deportivos',
    '27201': 'Prendas de protección personal',
    '27301': 'Artículos deportivos',
    '27401': 'Productos textiles',
    '27501': 'Blancos y otros productos textiles, excepto prendas de vestir',
    '28101': 'Sustancias y materiales explosivos',
    '28199': 'Materiales y suministros para seguridad',
    '28201': 'Materiales de seguridad pública',
    '28301': 'Prendas de protección para seguridad pública y nacional',
    '29101': 'Herramientas menores',
    '29199': 'Herramientas, refacciones y accesorios menores',
    '29201': 'Refacciones y accesorios menores de edificios',
    '29301': 'Refacciones y accesorios menores de mobiliario y equipo',
    '29401': 'Refacciones y accesorios para equipo de cómputo y telecomunicaciones',
    '29501': 'Refacciones y accesorios menores de equipo e instrumental médico',
    '29601': 'Refacciones y accesorios menores de equipo de transporte',
    '29701': 'Refacciones y accesorios menores de equipo de defensa y seguridad',
    '29801': 'Refacciones y accesorios menores de maquinaria y otros equipos',
    '29901': 'Refacciones y accesorios menores otros bienes muebles',
    '31101': 'Servicio de energía eléctrica',
    '31199': 'Servicios básicos',
    '31201': 'Servicio de gas',
    '31301': 'Servicio de agua',
    '31401': 'Servicio telefónico convencional',
    '31501': 'Servicio de telefonía celular',
    '31601': 'Servicio de radiolocalización',
    '31602': 'Servicios de telecomunicaciones',
    '31603': 'Servicios de Internet',
    '31701': 'Servicios de conducción de señales analógicas y digitales',
    '31801': 'Servicio postal',
    '31802': 'Servicio telegráfico',
    '31901': 'Servicios integrales de telecomunicación',
    '31902': 'Contratación de otros servicios',
    '31903': 'Servicios generales para planteles educativos',
    '31904': 'Servicios integrales de infraestructura de cómputo',
    '32101': 'Arrendamiento de terrenos',
    '32199': 'Servicios de arrendamiento',
    '32201': 'Arrendamiento de edificios y locales',
    '32301': 'Arrendamiento de equipo y bienes informáticos',
    '32302': 'Arrendamiento de mobiliario',
    '32303': 'Arrendamiento de equipo de telecomunicaciones',
    '32401': 'Arrendamiento de equipo e instrumental médico y de laboratorio',
    '32501': 'Arrendamiento de vehículos para seguridad pública',
    '32502': 'Arrendamiento de vehículos para servicios públicos',
    '32503': 'Arrendamiento de vehículos para servicios administrativos',
    '32504': 'Arrendamiento de vehículos para desastres naturales',
    '32505': 'Arrendamiento de vehículos para servidores públicos',
    '32601': 'Arrendamiento de maquinaria y equipo',
    '32701': 'Patentes, derechos de autor, regalías y otros',
    '32901': 'Arrendamiento de sustancias y productos químicos',
    '32902': 'PIDIREGAS cargos fijos',
    '32903': 'Otros arrendamientos',
    '33101': 'Asesorías asociadas a convenios, tratados o acuerdos',
    '33102': 'Asesorías por controversias en el marco de los tratados internacionales',
    '33103': 'Consultorías para programas o proyectos financiados por organismos internacionales',
    '33104': 'Otras asesorías para la operación de programas',
    '33105': 'Servicios relacionados con procedimientos jurisdiccionales',
    '33106': 'Servicios legales, de contabilidad, auditoría y relacionados',
    '33199': 'Servicios profesionales, científicos, técnicos y otros servicios',
    '33201': 'Servicios de diseño, arquitectura, ingeniería y actividades relacionadas',
    '33301': 'Servicios de desarrollo de aplicaciones informáticas',
    '33302': 'Servicios estadísticos y geográficos',
    '33303': 'Servicios relacionados con certificación de procesos',
    '33304': 'Servicios de mantenimiento de aplicaciones informáticas',
    '33401': 'Servicios para capacitación a servidores públicos',
    '33501': 'Estudios e investigaciones',
    '33601': 'Servicios relacionados con traducciones',
    '33602': 'Otros servicios comerciales',
    '33603': 'Impresiones de documentos oficiales',
    '33604': 'Impresión y elaboración de material informativo',
    '33605': 'Información en medios masivos',
    '33606': 'Servicios de digitalización',
    '33701': 'Gastos de seguridad pública y nacional',
    '33702': 'Gastos en actividades de seguridad y logística del Estado Mayor Presidencial',
    '33801': 'Servicios de vigilancia',
    '33901': 'Subcontratación de servicios con terceros',
    '33902': 'Proyectos para prestación de servicios',
    '33903': 'Servicios integrales',
    '33904': 'Asignaciones derivadas de proyectos de asociación público privada',
    '33905': 'Servicios integrales en materia de seguridad pública y nacional',
    '33906': 'Asignaciones para cubrir el pago de obligaciones derivadas de títulos de concesión',
    '34101': 'Servicios bancarios y financieros',
    '34199': 'Servicios financieros, bancarios y comerciales',
    '34301': 'Gastos inherentes a la recaudación',
    '34401': 'Seguro de responsabilidad patrimonial del Estado',
    '34501': 'Seguros de bienes patrimoniales',
    '34601': 'Almacenaje, embalaje y envase',
    '34701': 'Fletes y maniobras',
    '34801': 'Comisiones por ventas',
    '34901': 'Otros servicios financieros, bancarios y comerciales',
    '35101': 'Mantenimiento y conservación de inmuebles para servicios administrativos',
    '35102': 'Mantenimiento y conservación de inmuebles para servicios públicos',
    '35199': 'Servicios de instalación, reparación, mantenimiento y conservación',
    '35201': 'Mantenimiento y conservación de mobiliario y equipo de administración',
    '35301': 'Mantenimiento y conservación de bienes informáticos',
    '35401': 'Instalación, reparación y mantenimiento de equipo e instrumental médico',
    '35501': 'Mantenimiento y conservación de vehículos',
    '35601': 'Reparación y mantenimiento de equipo de defensa y seguridad',
    '35701': 'Mantenimiento y conservación de maquinaria y equipo',
    '35702': 'Mantenimiento y conservación de plantas e instalaciones productivas',
    '35801': 'Servicios de lavandería, limpieza e higiene',
    '35901': 'Servicios de jardinería y fumigación',
    '36101': 'Difusión de mensajes sobre programas y actividades gubernamentales',
    '36199': 'Servicios de comunicación social y publicidad',
    '36201': 'Difusión de mensajes comerciales para promover la venta de productos',
    '36301': 'Servicios de creatividad, preproducción y producción de publicidad',
    '36401': 'Servicios de revelado de fotografías',
    '36601': 'Servicio de creación y difusión de contenido a través de Internet',
    '36901': 'Servicios relacionados con monitoreo de información en medios masivos',
    '37101': 'Pasajes aéreos nacionales para labores en campo y de supervisión',
    '37102': 'Pasajes aéreos nacionales asociados a los programas de seguridad pública',
    '37103': 'Pasajes aéreos nacionales asociados a desastres naturales',
    '37104': 'Pasajes aéreos nacionales para servidores públicos de mando',
    '37105': 'Pasajes aéreos internacionales asociados a seguridad pública',
    '37106': 'Pasajes aéreos internacionales para servidores públicos',
    '37199': 'Servicios de traslado y viáticos',
    '37201': 'Pasajes terrestres nacionales para labores en campo',
    '37202': 'Pasajes terrestres nacionales asociados a seguridad pública',
    '37203': 'Pasajes terrestres nacionales asociados a desastres naturales',
    '37204': 'Pasajes terrestres nacionales para servidores públicos de mando',
    '37205': 'Pasajes terrestres internacionales asociados a seguridad pública',
    '37206': 'Pasajes terrestres internacionales para servidores públicos',
    '37207': 'Pasajes terrestres nacionales por medio electrónico',
    '37301': 'Pasajes marítimos para labores en campo y de supervisión',
    '37302': 'Pasajes marítimos asociados a seguridad pública',
    '37303': 'Pasajes marítimos asociados a desastres naturales',
    '37304': 'Pasajes marítimos para servidores públicos de mando',
    '37501': 'Viáticos nacionales para labores en campo y de supervisión',
    '37502': 'Viáticos nacionales asociados a seguridad pública',
    '37503': 'Viáticos nacionales asociados a desastres naturales',
    '37504': 'Viáticos nacionales para servidores públicos',
    '37601': 'Viáticos en el extranjero asociados a seguridad pública',
    '37602': 'Viáticos en el extranjero para servidores públicos',
    '37701': 'Instalación del personal federal',
    '37801': 'Servicios integrales nacionales para servidores públicos',
    '37802': 'Servicios integrales en el extranjero para servidores públicos',
    '37901': 'Gastos para operativos y trabajos de campo en áreas rurales',
    '38101': 'Gastos de ceremonial del titular del Ejecutivo Federal',
    '38102': 'Gastos de ceremonial de los titulares de las dependencias',
    '38103': 'Gastos inherentes a la investidura presidencial',
    '38199': 'Servicios oficiales',
    '38201': 'Gastos de orden social',
    '38301': 'Congresos y convenciones',
    '38401': 'Exposiciones',
    '38501': 'Gastos para alimentación de servidores públicos de mando',
    '39101': 'Funerales y pagas de defunción',
    '39199': 'Otros servicios generales',
    '39201': 'Impuestos y derechos de exportación',
    '39202': 'Otros impuestos y derechos',
    '39301': 'Impuestos y derechos de importación',
    '39401': 'Erogaciones por resoluciones por autoridad competente',
    '39402': 'Indemnizaciones por expropiación de predios',
    '39403': 'Otras asignaciones derivadas de resoluciones de ley',
    '39501': 'Penas, multas, accesorios y actualizaciones',
    '39601': 'Pérdidas del erario federal',
    '39602': 'Otros gastos por responsabilidades',
    '39701': 'Erogaciones por pago de utilidades',
    '39801': 'Impuesto sobre nóminas',
    '39810': 'Otros impuestos sobre nóminas',
    '39901': 'Gastos de las Comisiones Internacionales de Límites y Aguas',
    '39902': 'Gastos de las oficinas del Servicio Exterior Mexicano',
    '39903': 'Asignaciones a los grupos parlamentarios',
    '39904': 'Participaciones en órganos de gobierno',
    '39905': 'Actividades de coordinación con el Presidente Electo',
    '39906': 'Servicios Corporativos prestados por las Entidades Paraestatales',
    '39907': 'Servicios prestados entre Organismos de una Entidad Paraestatal',
    '39908': 'Erogaciones por cuenta de terceros',
    '39909': 'Erogaciones recuperables',
    '39910': 'Apertura de Fondo Rotatorio',
    '41501': 'Transferencias para cubrir el déficit de operación',
    '41601': 'Transferencias a entidades empresariales no financieras',
    '43101': 'Subsidios a la producción',
    '43201': 'Subsidios a la distribución',
    '43301': 'Subsidios para inversión',
    '43401': 'Subsidios a la prestación de servicios públicos',
    '43501': 'Subsidios para cubrir diferenciales de tasas de interés',
    '43601': 'Subsidios para la adquisición de vivienda de interés social',
    '43701': 'Subsidios al consumo',
    '43801': 'Subsidios a Entidades Federativas y Municipios',
    '43901': 'Subsidios para capacitación y becas',
    '43902': 'Subsidios a fideicomisos privados y estatales',
    '44101': 'Gastos relacionados con actividades culturales, deportivas y de ayuda extraordinaria',
    '44102': 'Gastos por servicios de traslado de personas',
    '44103': 'Premios, recompensas, pensiones de gracia y pensión recreativa estudiantil',
    '44104': 'Premios, estímulos, recompensas, becas y seguros a deportistas',
    '44105': 'Apoyo a voluntarios que participan en diversos programas federales',
    '44106': 'Compensaciones por servicios de carácter social',
    '44199': 'Ayudas sociales',
    '44201': 'Otras ayudas para programas de capacitación',
    '44401': 'Apoyos a la investigación científica y tecnológica',
    '44402': 'Apoyos a la investigación científica en instituciones sin fines de lucro',
    '44801': 'Mercancías para su distribución a la población',
    '45201': 'Pago de pensiones y jubilaciones',
    '45202': 'Pago de pensiones y jubilaciones contractuales',
    '45203': 'Transferencias para el pago de pensiones y jubilaciones',
    '45901': 'Pago de sumas aseguradas',
    '45902': 'Prestaciones económicas distintas de pensiones y jubilaciones',
    '46101': 'Aportaciones a fideicomisos públicos',
    '46102': 'Aportaciones a mandatos públicos',
    '46199': 'Transferencias a fideicomisos, mandatos y otros análogos',
    '47101': 'Trasferencias para cuotas y aportaciones de seguridad social',
    '47102': 'Transferencias para cuotas y aportaciones a los seguros de retiro',
    '48101': 'Donativos a instituciones sin fines de lucro',
    '48199': 'Donativos',
    '48201': 'Donativos a entidades federativas o municipios',
    '48301': 'Donativos a fideicomisos privados',
    '48401': 'Donativos a fideicomisos estatales',
    '48501': 'Donativos internacionales',
    '49199': 'Transferencias al exterior',
    '49201': 'Cuotas y aportaciones a organismos internacionales',
    '49202': 'Otras aportaciones internacionales',
    '51101': 'Mobiliario',
    '51199': 'Mobiliario y equipo de administración',
    '51201': 'Muebles, excepto de oficina y estantería',
    '51301': 'Bienes artísticos y culturales',
    '51501': 'Bienes informáticos',
    '51901': 'Equipo de administración',
    '51902': 'Adjudicaciones, expropiaciones e indemnizaciones de bienes muebles',
    '52101': 'Equipos y aparatos audiovisuales',
    '52199': 'Mobiliario y equipo educacional y recreativo',
    '52201': 'Aparatos deportivos',
    '52301': 'Cámaras fotográficas y de video',
    '52901': 'Otro mobiliario y equipo educacional y recreativo',
    '53101': 'Equipo médico y de laboratorio',
    '53199': 'Equipo e instrumental médico y de laboratorio',
    '53201': 'Instrumental médico y de laboratorio',
    '54101': 'Vehículos y equipo terrestres para seguridad pública',
    '54102': 'Vehículos y equipo terrestres para desastres naturales',
    '54103': 'Vehículos y equipo terrestres para servicios públicos',
    '54104': 'Vehículos y equipo terrestres para servicios administrativos',
    '54105': 'Vehículos y equipo terrestres para servidores públicos',
    '54199': 'Vehículos y equipo de transporte',
    '54201': 'Carrocerías y remolques',
    '54301': 'Vehículos y equipo aéreos para seguridad pública',
    '54302': 'Vehículos y equipo aéreos para desastres naturales',
    '54303': 'Vehículos y equipo aéreos para servicios públicos',
    '54401': 'Equipo ferroviario',
    '54501': 'Vehículos y equipo marítimo para seguridad pública',
    '54502': 'Vehículos y equipo marítimo para servicios públicos',
    '54503': 'Construcción de embarcaciones',
    '54901': 'Otros equipos de transporte',
    '55101': 'Maquinaria y equipo de defensa y seguridad pública',
    '55102': 'Equipo de seguridad pública y nacional',
    '55199': 'Equipo de defensa y seguridad',
    '56101': 'Maquinaria y equipo agropecuario',
    '56199': 'Maquinaria, otros equipos y herramientas',
    '56201': 'Maquinaria y equipo industrial',
    '56301': 'Maquinaria y equipo de construcción',
    '56401': 'Sistemas de aire acondicionado, calefacción y de refrigeración',
    '56501': 'Equipos y aparatos de comunicaciones y telecomunicaciones',
    '56601': 'Maquinaria y equipo eléctrico y electrónico',
    '56701': 'Herramientas y máquinas herramienta',
    '56901': 'Bienes muebles por arrendamiento financiero',
    '56902': 'Otros bienes muebles',
    '57101': 'Animales de reproducción',
    '57199': 'Activos biológicos',
    '57201': 'Porcinos',
    '57301': 'Aves',
    '57401': 'Ovinos y caprinos',
    '57501': 'Peces y acuicultura',
    '57601': 'Animales de trabajo',
    '57701': 'Animales de custodia y vigilancia',
    '57801': 'Árboles y plantas',
    '57901': 'Otros activos biológicos',
    '58101': 'Terrenos',
    '58199': 'Bienes inmuebles',
    '58301': 'Edificios y locales',
    '58901': 'Adjudicaciones, expropiaciones e indemnizaciones de inmuebles',
    '58902': 'Bienes inmuebles en la modalidad de proyectos de infraestructura',
    '58903': 'Bienes inmuebles por arrendamiento financiero',
    '58904': 'Otros bienes inmuebles',
    '59101': 'Software',
    '59199': 'Activos intangibles',
    '59401': 'Derechos',
    '59701': 'Licencias informáticas e intelectuales',
    '59901': 'Otros activos intangibles',
}

CATALOGO_PROGRAMAS = {
    'B004': 'Adquisición de leche nacional',
    'B005': 'Producción y comercialización de Biológicos Veterinarios',
    'B006': 'Adquisición, industrialización y comercialización de productos agroalimentarios',
    'E001': 'Desarrollo, aplicación de programas educativos e investigación en materia agroalimentaria',
    'E006': 'Generación de Proyectos de Investigación',
    'G001': 'Regulación, supervisión y aplicación de las políticas públicas',
    'K017': 'Infraestructura para el desarrollo rural sustentable',
    'M001': 'Actividades de apoyo administrativo',
    'O001': 'Actividades de apoyo a la función pública y buen gobierno',
    'P001': 'Diseño y Aplicación de la Política Agropecuaria',
    'P021': 'Aplicación de la Política Agropecuaria',
    'Q004': 'Desarrollo y aplicación de programas y proyectos educativos y de investigación en el sector agroalimentario',
    'S052': 'Programa de Abasto Social y Precios de Garantía a cargo de Leche para el Bienestar, S.A. de C.V.',
    'S053': 'Programa de Abasto Rural',
    'S263': 'Sanidad e Inocuidad Agroalimentaria',
    'S290': 'Acopio para el Bienestar',
    'S292': 'Fertilizantes para el Bienestar',
    'S293': 'Producción para el Bienestar',
    'S304': 'Pesca y Acuacultura Sustentables',
    'S318': 'Comercio Justo',
    'U027': 'Soberanía Alimentaria',
    'W001': 'Operaciones ajenas',
}

CATALOGO_CAPITULOS = {
    '1000': 'Servicios Personales',
    '2000': 'Materiales y Suministros',
    '3000': 'Servicios Generales',
    '4000': 'Transferencias, Asignaciones, Subsidios y Otras Ayudas',
    '5000': 'Bienes Muebles, Inmuebles e Intangibles',
    '6000': 'Inversión Pública',
    '7000': 'Inversiones Financieras y Otras Provisiones',
    '8000': 'Participaciones y Aportaciones',
    '9000': 'Deuda Pública',
}

# ---------------------------------------------------------------------------
# Definición de familias de importes por fuente
# (columna total ya existente en el crudo, o None si hay que sumarla)
# ---------------------------------------------------------------------------
def _cols_sicop(prefijo_2letras: list[str]) -> list[str]:
    return prefijo_2letras


FAMILIAS_SICOP = {
    "ORIGINAL": (["OREN", "ORFE", "ORMR", "ORAB", "ORMY", "ORJN", "ORJL", "ORAG", "ORSE", "OROC", "ORNO", "ORDI"], "ORIGINAL"),
    "AMPLIACION": (["AMEN", "AMFE", "AMMR", "AMAB", "AMMY", "AMJN", "AMJL", "AMAG", "AMSE", "AMOC", "AMNO", "AMDI"], "AMPLIACION"),
    "REDUCCION": (["REEN", "REFE", "REMR", "REAB", "REMY", "REJN", "REJL", "REAG", "RESE", "REOC", "RENO", "REDI"], "REDUCCION"),
    "RESERVAS": (["RESERVA_ENE", "RESERVA_FEB", "RESERVA_MZO", "RESERVA_ABR", "RESERVA_MAY", "RESERVA_JUN",
                  "RESERVA_JUL", "RESERVA_AGO", "RESERVA_SEP", "RESERVA_OCT", "RESERVA_NOV", "RESERVA_DIC"], "RESERVAS"),
    "MODIFICADO_AUTORIZADO": (["MOEN", "MOFE", "MOMR", "MOAB", "MOMY", "MOJN", "MOJL", "MOAG", "MOSE", "MOOC", "MONO", "MODI"], "MODIFICADO_AUTORIZADO"),
    "COMPROMETIDO": (["COEN", "COFE", "COMR", "COAB", "COMY", "COJN", "COJL", "COAG", "COSE", "COOC", "CONO", "CODI"], "COMPROMETIDO"),
    "EJERCIDO": (["EJEN", "EJFE", "EJMR", "EJAB", "EJMY", "EJJN", "EJJL", "EJAG", "EJSE", "EJOC", "EJNO", "EJDI"], "EJERCIDO"),
    "DEVENGADO": (["DVEN", "DVFE", "DVMR", "DVAB", "DVMY", "DVJN", "DVJL", "DVAG", "DVSE", "DVOC", "DVNO", "DVDI"], "DEVENGADO"),
    "EJERCIDO_TRAMITE": (["EJTREN", "EJTRFE", "EJTRMR", "EJTRAB", "EJTRMY", "EJTRJN", "EJTRJL", "EJTRAG",
                          "EJTRSE", "EJTROC", "EJTRNO", "EJTRDI"], "EJERCIDO_TRAMITE"),
}
CLAVES_SICOP = {"modificado": "MODIFICADO_AUTORIZADO", "ejercido": "EJERCIDO", "comprometido": "COMPROMETIDO"}

FAMILIAS_MAP = {
    f"{pref}": ([f"{pref}_{m}" for m in MESES], None)
    for pref in ["ORI", "AMP", "RED", "MOD", "CONG", "DESCONG", "EJE"]
}
CLAVES_MAP = {"modificado": "MOD", "ejercido": "EJE", "comprometido": None}

NOMBRES_FAMILIA = {
    "ORIGINAL": "Original", "ORI": "Original",
    "AMPLIACION": "Ampliación", "AMP": "Ampliación",
    "REDUCCION": "Reducción", "RED": "Reducción",
    "RESERVAS": "Reservas",
    "MODIFICADO_AUTORIZADO": "Modificado", "MOD": "Modificado",
    "COMPROMETIDO": "Comprometido",
    "EJERCIDO": "Ejercido", "EJE": "Ejercido",
    "DEVENGADO": "Devengado",
    "EJERCIDO_TRAMITE": "Ejercido en trámite",
    "CONG": "Congelado", "DESCONG": "Descongelado",
}

FUENTES = {
    "MAP": "Módulo de Adecuaciones Presupuestarias (MAP)",
    "SICOP": "Sistema de Contabilidad y Presupuesto (SICOP)",
}

# Normalización de Unidad Responsable — tomada tal cual de MAPEO_UR_2026_BASE y
# FUSION_URS_2026 en nrbeca/nuevo/config.py, para que claves legadas o alternas
# se agrupen bajo el código vigente antes de buscar el nombre.
MAPEO_UR_BASE = {
    "G00": "811", "108": "810", "113": "250", "121": "260", "122": "261", "123": "262",
    "124": "263", "125": "264", "126": "265", "127": "266", "128": "267", "129": "268",
    "130": "269", "131": "270", "132": "271", "133": "272", "134": "273", "135": "274",
    "136": "275", "137": "276", "138": "277", "139": "278", "140": "279", "141": "280",
    "142": "281", "143": "282", "144": "283", "145": "284", "146": "285", "147": "286",
    "148": "287", "149": "288", "150": "289", "151": "290", "152": "291", "153": "292",
    "215": "220", "300": "225", "310": "226", "700": "227", "600": "230", "612": "231",
    "312": "232", "315": "233", "400": "235", "311": "237", "314": "245",
}
FUSION_URS = {
    "810": "119", "812": "119", "800": "120", "811": "120", "235": "250", "236": "253",
    "237": "253", "225": "900", "245": "910", "241": "911", "246": "912", "247": "912",
    "230": "920", "226": "921", "227": "922", "231": "923", "232": "924",
}


def normalizar_ur(codigo) -> str:
    """Aplica el mismo encadenamiento MAPEO_UR_2026_BASE -> FUSION_URS_2026 que
    usa tu Dashboard de Austeridad, para que un código legado o alterno caiga
    bajo el mismo código vigente que usa el catálogo de nombres."""
    clave = str(codigo).strip()
    clave = MAPEO_UR_BASE.get(clave, clave)
    clave = FUSION_URS.get(clave, clave)
    return clave


# ---------------------------------------------------------------------------
# Catálogos (código -> nombre) — usan los diccionarios embebidos de arriba.
# Si algún día quieres editar un catálogo sin tocar el código, puedes seguir
# poniendo un catalogs/<nombre_archivo> junto al script: si existe, sus
# filas se agregan encima del catálogo embebido (y lo pisan si el código se
# repite).
# ---------------------------------------------------------------------------
CATALOGS_DIR = Path(__file__).parent / "catalogs"
_CATALOGOS_EMBEBIDOS = {
    "unidades.csv": CATALOGO_UNIDADES,
    "partidas.csv": CATALOGO_PARTIDAS,
    "programas.csv": CATALOGO_PROGRAMAS,
    "capitulos.csv": CATALOGO_CAPITULOS,
}


@st.cache_data(show_spinner=False)
def cargar_catalogo(nombre_archivo: str, col_codigo: str, col_nombre: str) -> dict[str, str]:
    cat = dict(_CATALOGOS_EMBEBIDOS.get(nombre_archivo, {}))
    ruta = CATALOGS_DIR / nombre_archivo
    if ruta.exists():
        try:
            extra = pd.read_csv(ruta, dtype=str)
            extra[col_codigo] = extra[col_codigo].str.strip()
            cat.update(dict(zip(extra[col_codigo], extra[col_nombre])))
        except Exception:
            pass
    return cat


def etiqueta_con_nombre(codigo, catalogo: dict[str, str]) -> str:
    if pd.isna(codigo):
        return "(sin dato)"
    clave = str(codigo).strip()
    clave_num = clave.split(".")[0] if clave.replace(".", "", 1).isdigit() else clave
    nombre = catalogo.get(clave) or catalogo.get(clave_num)
    return f"{clave} — {nombre}" if nombre else clave


# ---------------------------------------------------------------------------
# Carga y preparación de datos crudos
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner="Leyendo archivo...")
def cargar_crudo(archivo_bytes: bytes, nombre_archivo: str, fuente: str) -> pd.DataFrame:
    buffer = io.BytesIO(archivo_bytes)
    if nombre_archivo.lower().endswith((".xlsx", ".xls")):
        df = pd.read_excel(buffer)
    else:
        encodings = ["utf-8", "latin-1", "cp1252"]
        df = None
        for enc in encodings:
            try:
                buffer.seek(0)
                df = pd.read_csv(buffer, encoding=enc, low_memory=False)
                break
            except UnicodeDecodeError:
                continue
        if df is None:
            raise ValueError("No se pudo leer el archivo con ninguna codificación común (utf-8/latin-1/cp1252).")
    df.columns = [str(c).strip() for c in df.columns]

    if fuente == "SICOP":
        for c in ["CAPITULO", "CONCEPTO", "PARTIDA_GENERICA", "PARTIDA_ESPECIFICA"]:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)
        if {"CAPITULO", "CONCEPTO", "PARTIDA_GENERICA", "PARTIDA_ESPECIFICA"}.issubset(df.columns):
            df["Partida"] = (df["CAPITULO"] * 10000 + df["CONCEPTO"] * 1000
                              + df["PARTIDA_GENERICA"] * 100 + df["PARTIDA_ESPECIFICA"])
        if "ID_UNIDAD" in df.columns:
            df["Unidad Responsable (código en la base)"] = df["ID_UNIDAD"].astype(str).str.strip()
            df["Unidad Responsable"] = df["Unidad Responsable (código en la base)"].apply(normalizar_ur)
        if "PROGRAMA_PRESUPUESTARIO" in df.columns:
            df["Programa"] = df["PROGRAMA_PRESUPUESTARIO"].astype(str).str.strip()
    elif fuente == "MAP":
        if "PARTIDA" in df.columns:
            df["Partida"] = pd.to_numeric(df["PARTIDA"], errors="coerce")
        if "UNIDAD" in df.columns:
            df["Unidad Responsable (código en la base)"] = df["UNIDAD"].astype(str).str.strip()
            df["Unidad Responsable"] = df["Unidad Responsable (código en la base)"].apply(normalizar_ur)
        if "PROGRAMA" in df.columns:
            df["Programa"] = df["PROGRAMA"].astype(str).str.strip()

    if "Partida" in df.columns:
        df["Capitulo"] = (pd.to_numeric(df["Partida"], errors="coerce") // 10000).astype("Int64")

    return df


def solo_nombre(codigo, catalogo: dict[str, str]) -> str:
    if pd.isna(codigo):
        return ""
    clave = str(codigo).strip()
    clave_num = clave.split(".")[0] if clave.replace(".", "", 1).isdigit() else clave
    return catalogo.get(clave) or catalogo.get(clave_num) or ""


def enriquecer_con_catalogos(df: pd.DataFrame) -> pd.DataFrame:
    cat_ur = cargar_catalogo("unidades.csv", "codigo_ur", "nombre_ur")
    cat_partidas = cargar_catalogo("partidas.csv", "partida", "nombre_partida")
    cat_programas = cargar_catalogo("programas.csv", "programa", "nombre_programa")
    cat_capitulos = cargar_catalogo("capitulos.csv", "capitulo", "nombre_capitulo")

    nuevas = {}
    if "Unidad Responsable" in df.columns:
        nuevas["Nombre de la Unidad Responsable"] = df["Unidad Responsable"].apply(lambda x: solo_nombre(x, cat_ur))
        nuevas["Unidad Responsable (nombre)"] = df["Unidad Responsable"].apply(lambda x: etiqueta_con_nombre(x, cat_ur))
    if "Partida" in df.columns:
        nuevas["Nombre Partida"] = df["Partida"].apply(lambda x: solo_nombre(x, cat_partidas))
        nuevas["Partida (nombre)"] = df["Partida"].apply(lambda x: etiqueta_con_nombre(x, cat_partidas))
    if "Programa" in df.columns:
        nuevas["Nombre Programa"] = df["Programa"].apply(lambda x: solo_nombre(x, cat_programas))
        nuevas["Programa (nombre)"] = df["Programa"].apply(lambda x: etiqueta_con_nombre(x, cat_programas))
    if "Capitulo" in df.columns:
        nuevas["Nombre Capítulo"] = df["Capitulo"].apply(lambda x: solo_nombre(x, cat_capitulos))
        nuevas["Capítulo (nombre)"] = df["Capitulo"].apply(lambda x: etiqueta_con_nombre(x, cat_capitulos))
    return pd.concat([df, pd.DataFrame(nuevas, index=df.index)], axis=1)


def agregar_periodos_y_disponible(df: pd.DataFrame, fuente: str, mes_corte_idx: int) -> tuple[pd.DataFrame, list[str]]:
    """Agrega columnas '<Familia> (Anual)' y '<Familia> (Al <mes>)' por cada
    familia de importes disponible, más 'Disponible (Anual)'/'Disponible (Al <mes>)'.
    Regresa el df enriquecido y la lista de columnas de valor calculadas (en orden lógico)."""
    familias = FAMILIAS_SICOP if fuente == "SICOP" else FAMILIAS_MAP
    claves = CLAVES_SICOP if fuente == "SICOP" else CLAVES_MAP
    mes_label = NOMBRES_MES[mes_corte_idx].capitalize()

    df = df.copy()
    columnas_valor = []
    for fam, (cols_mensuales, col_total_existente) in familias.items():
        presentes = [c for c in cols_mensuales if c in df.columns]
        if not presentes:
            continue
        nombre_bonito = NOMBRES_FAMILIA.get(fam, fam)
        col_anual = f"{nombre_bonito} (Anual)"
        if col_total_existente and col_total_existente in df.columns:
            df[col_anual] = pd.to_numeric(df[col_total_existente], errors="coerce").fillna(0)
        else:
            df[col_anual] = df[presentes].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)

        cols_al_periodo = [c for c in cols_mensuales[: mes_corte_idx + 1] if c in df.columns]
        col_periodo = f"{nombre_bonito} (Al {mes_label})"
        df[col_periodo] = df[cols_al_periodo].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1) if cols_al_periodo else 0.0

        columnas_valor += [col_anual, col_periodo]

    # Disponible = Modificado - Ejercido - Comprometido (verificado contra el
    # archivo de ejemplo: cuadra al centavo con las columnas "Importe Disponible").
    nombre_mod = NOMBRES_FAMILIA.get(claves["modificado"])
    nombre_eje = NOMBRES_FAMILIA.get(claves["ejercido"])
    nombre_com = NOMBRES_FAMILIA.get(claves["comprometido"]) if claves["comprometido"] else None

    col_mod_anual, col_mod_periodo = f"{nombre_mod} (Anual)", f"{nombre_mod} (Al {mes_label})"
    col_eje_anual = f"{nombre_eje} (Anual)"
    if col_mod_anual in df.columns and col_eje_anual in df.columns:
        com_anual = df[f"{nombre_com} (Anual)"] if nombre_com and f"{nombre_com} (Anual)" in df.columns else 0
        com_periodo = df[f"{nombre_com} (Al {mes_label})"] if nombre_com and f"{nombre_com} (Al {mes_label})" in df.columns else 0
        df["Disponible (Anual)"] = df[col_mod_anual] - df[col_eje_anual] - com_anual
        df["Disponible (Al " + mes_label + ")"] = df[col_mod_periodo] - df[col_eje_anual] - com_periodo
        columnas_valor += ["Disponible (Anual)", f"Disponible (Al {mes_label})"]

    # "Ejercido real" (solo SICOP) = Ejercido + Devengado + Ejercido en trámite,
    # tal como lo define tu Dashboard de Presupuesto (EJERCIDO_REAL en
    # sicop_processor.py). Se agrega aparte, sin tocar "Ejercido", porque el
    # formato estándar de Estado del Ejercicio OREF usa Ejercido solo.
    if fuente == "SICOP" and all(f"{n} (Anual)" in df.columns for n in ["Ejercido", "Devengado", "Ejercido en trámite"]):
        df["Ejercido real (Anual)"] = df["Ejercido (Anual)"] + df["Devengado (Anual)"] + df["Ejercido en trámite (Anual)"]
        df[f"Ejercido real (Al {mes_label})"] = (df["Ejercido (Anual)"] + df[f"Devengado (Al {mes_label})"]
                                                  + df[f"Ejercido en trámite (Al {mes_label})"])
        columnas_valor += ["Ejercido real (Anual)", f"Ejercido real (Al {mes_label})"]

    return df, columnas_valor


def construir_reporte_plantilla(df: pd.DataFrame, fuente: str, mes_corte_idx: int):
    """Reproduce exactamente el formato de formato_estado_del_ejercicio.xlsx:
    Unidad Responsable, Nombre UR, Partida, Nombre Partida + Autorizado/
    Reservado/Modificado/Comprometido (Anual y Al periodo) + Ejercido +
    Disponible (Anual y Al periodo). Si la base no trae alguno de estos
    conceptos (ej. MAP no tiene Reservado ni Comprometido), esa columna
    simplemente se omite. Regresa (tabla, encabezados, grupos, filas, roles)."""
    mes_label = NOMBRES_MES[mes_corte_idx].capitalize()
    filas = [c for c in ["Unidad Responsable", "Nombre de la Unidad Responsable", "Partida", "Nombre Partida"] if c in df.columns]

    pares = [("Original", "Importe Autorizado", "autorizado"), ("Reservas", "Importe Reservado", "reservado"),
             ("Modificado", "Importe Modificado", "modificado"), ("Comprometido", "Importe Comprometido", "comprometido")]

    especificacion = []  # (columna_interna, encabezado, grupo, rol)
    for base, etiqueta, rol in pares:
        col = f"{base} (Anual)"
        if col in df.columns:
            especificacion.append((col, etiqueta, "Anual", rol))
    for base, etiqueta, rol in pares:
        col = f"{base} (Al {mes_label})"
        if col in df.columns:
            especificacion.append((col, etiqueta, "Al periodo", rol))
    if "Ejercido (Anual)" in df.columns:
        especificacion.append(("Ejercido (Anual)", "Importe Ejercido", None, "ejercido"))
    if "Disponible (Anual)" in df.columns:
        especificacion.append(("Disponible (Anual)", "Importe Disponible", "Anual", "disponible"))
    if f"Disponible (Al {mes_label})" in df.columns:
        especificacion.append((f"Disponible (Al {mes_label})", "Importe Disponible", "Al periodo", "disponible"))

    cols_internas = [c for c, _, _, _ in especificacion]
    if not filas or not cols_internas:
        return pd.DataFrame(), [], [], filas, []

    agregado = df.groupby(filas, as_index=False)[cols_internas].sum()
    fila_total = {c: "" for c in agregado.columns}
    fila_total[filas[0]] = "Total general"
    for c in cols_internas:
        fila_total[c] = agregado[c].sum()
    agregado = pd.concat([pd.DataFrame([fila_total]), agregado], ignore_index=True)
    agregado = agregado[filas + cols_internas]

    encabezados = filas + [etiqueta for _, etiqueta, _, _ in especificacion]

    grupos = []
    n_filas = len(filas)
    i = 0
    while i < len(especificacion):
        grupo = especificacion[i][2]
        if grupo in ("Anual", "Al periodo"):
            j = i
            while j < len(especificacion) and especificacion[j][2] == grupo:
                j += 1
            texto = "Anual" if grupo == "Anual" else "Al periodo"
            grupos.append((n_filas + i + 1, n_filas + j, texto))
            i = j
        else:
            i += 1

    # roles: lista paralela a las columnas de valor -> (rol, periodo) para
    # que el exportador pueda escribir fórmulas de Excel reales (Disponible
    # = Modificado - Ejercido - Comprometido) en vez de valores fijos.
    roles = [(rol, "anual" if grupo == "Anual" else ("periodo" if grupo == "Al periodo" else None))
             for _, _, grupo, rol in especificacion]

    return agregado, encabezados, grupos, filas, roles


def aplicar_depuracion_sicop(df: pd.DataFrame) -> pd.DataFrame:
    """Reglas confirmadas en sicop_processor.py (nrbeca/nuevo) para el
    Estado del Ejercicio: excluir capítulo 1000 (servicios personales),
    la partida 39801, y CONTROL_OPERATIVO entre 60 y 69. Revisa que
    correspondan al reporte que quieres armar antes de activarlas."""
    dff = df.copy()
    col_cap = "Capitulo" if "Capitulo" in dff.columns else ("CAPITULO" if "CAPITULO" in dff.columns else None)
    if col_cap:
        dff = dff[dff[col_cap] != 1]
    if "Partida" in dff.columns:
        dff = dff[dff["Partida"] != 39801]
    if "CONTROL_OPERATIVO" in dff.columns:
        co = pd.to_numeric(dff["CONTROL_OPERATIVO"], errors="coerce")
        dff = dff[~co.between(60, 69)]
    return dff


def columnas_categoricas(df: pd.DataFrame, columnas_valor: list[str]) -> list[str]:
    return [c for c in df.columns if c not in columnas_valor and not pd.api.types.is_float_dtype(df[c])] or \
        [c for c in df.columns if c not in columnas_valor]


# ---------------------------------------------------------------------------
# Construcción de la tabla dinámica
# ---------------------------------------------------------------------------
def construir_pivote(df, filas, columnas, valores, filtros) -> pd.DataFrame:
    dff = df.copy()
    for col, seleccion in filtros.items():
        if seleccion:
            dff = dff[dff[col].isin(seleccion)]

    if not filas or not valores:
        return pd.DataFrame()

    if columnas:
        pivote = pd.pivot_table(dff, index=filas, columns=columnas, values=valores, aggfunc="sum", fill_value=0)
        pivote.columns = [" | ".join(str(x) for x in c) if isinstance(c, tuple) else str(c) for c in pivote.columns]
        pivote = pivote.reset_index()
    else:
        pivote = dff.groupby(filas, as_index=False)[valores].sum()

    fila_total = {c: "" for c in pivote.columns}
    if filas:
        fila_total[filas[0]] = "Total general"
    for c in pivote.columns:
        if c not in filas and pd.api.types.is_numeric_dtype(pivote[c]):
            fila_total[c] = pivote[c].sum()
    pivote = pd.concat([pd.DataFrame([fila_total]), pivote], ignore_index=True)
    return pivote


# ---------------------------------------------------------------------------
# Exportación a Excel con el formato del Estado del Ejercicio
# ---------------------------------------------------------------------------
def exportar_excel_oref(pivote: pd.DataFrame, fuente: str, linea1: str, linea2: str,
                         titulo: str, subtitulo: str, filas: list[str],
                         encabezados: list[str] | None = None,
                         grupos: list[tuple[int, int, str]] | None = None,
                         roles: list[tuple[str, str | None]] | None = None) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Reporte"
    ws.sheet_view.showGridLines = False

    n_cols = max(len(pivote.columns), 1)
    ultima_col = get_column_letter(n_cols)
    encabezados = encabezados or [str(c) for c in pivote.columns]

    c = ws.cell(row=1, column=n_cols, value=linea1)
    c.font = Font(name="Arial", size=12, bold=True)
    c.alignment = Alignment(horizontal="right")

    c = ws.cell(row=2, column=n_cols, value=linea2)
    c.font = Font(name="Arial", size=11, bold=True)
    c.alignment = Alignment(horizontal="right")

    ws.merge_cells(f"A5:{ultima_col}5")
    c = ws["A5"]; c.value = titulo
    c.font = Font(name="Calibri", size=14, bold=True)
    c.alignment = Alignment(horizontal="center")
    ws.row_dimensions[5].height = 18.75

    ws.merge_cells(f"A6:{ultima_col}6")
    c = ws["A6"]; c.value = subtitulo
    c.font = Font(name="Calibri", size=12, italic=True)
    c.alignment = Alignment(horizontal="center")
    ws.row_dimensions[6].height = 15.75

    fila_grupo = 8
    fila_encabezado = 9
    for col_ini, col_fin, texto in (grupos or []):
        if col_fin > col_ini:
            ws.merge_cells(start_row=fila_grupo, start_column=col_ini, end_row=fila_grupo, end_column=col_fin)
        celda = ws.cell(row=fila_grupo, column=col_ini, value=texto)
        celda.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor=VERDE)
        celda.alignment = Alignment(horizontal="center", vertical="top")
        for j in range(col_ini, col_fin + 1):
            ws.cell(row=fila_grupo, column=j).border = Border(left=THIN_BLANCO, right=THIN_BLANCO)

    ws.row_dimensions[fila_encabezado].height = 30
    for j, texto in enumerate(encabezados, start=1):
        celda = ws.cell(row=fila_encabezado, column=j, value=texto)
        celda.font = Font(name="Arial", size=11, bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor=BURDEOS)
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        celda.border = BORDE_ENCABEZADO

    n_filas_agrupadoras = max(len(filas), 1)

    # Mapa (rol, periodo) -> letra de columna, para poder escribir fórmulas
    # reales (Disponible = Modificado - Ejercido - Comprometido) en vez de
    # valores fijos.
    col_por_rol = {}
    if roles:
        for idx, (rol, periodo) in enumerate(roles):
            col_por_rol[(rol, periodo)] = get_column_letter(n_filas_agrupadoras + 1 + idx)

    total_row_num = fila_encabezado + 1
    first_data_row = total_row_num + 1
    last_row_num = fila_encabezado + len(pivote)

    for i, (_, fila) in enumerate(pivote.iterrows()):
        r = fila_encabezado + 1 + i
        es_total = str(fila.iloc[0]) == "Total general"
        for j, col in enumerate(pivote.columns, start=1):
            valor = fila[col]
            es_columna_valor = col not in (filas or [])
            idx_rol = j - n_filas_agrupadoras - 1
            rol, periodo = roles[idx_rol] if (roles and es_columna_valor and 0 <= idx_rol < len(roles)) else (None, None)

            celda = ws.cell(row=r, column=j)
            if rol == "disponible":
                col_mod = col_por_rol.get(("modificado", periodo))
                col_eje = col_por_rol.get(("ejercido", None))
                col_com = col_por_rol.get(("comprometido", periodo))
                if col_mod and col_eje:
                    formula = f"={col_mod}{r}-{col_eje}{r}"
                    if col_com:
                        formula += f"-{col_com}{r}"
                    celda.value = formula
                else:
                    celda.value = valor
            elif es_total and es_columna_valor and last_row_num >= first_data_row:
                col_letra = get_column_letter(j)
                celda.value = f"=SUM({col_letra}{first_data_row}:{col_letra}{last_row_num})"
            else:
                celda.value = valor

            celda.border = BORDE
            celda.alignment = Alignment(vertical="center", horizontal="center" if col in (filas or []) else None)
            if es_columna_valor and (isinstance(valor, (int, float)) or rol):
                celda.number_format = FORMATO_MONEDA
            if es_total:
                celda.font = Font(name="Arial", size=11, bold=True)
                celda.fill = PatternFill("solid", fgColor=GRIS_TOTAL)
            else:
                celda.font = Font(name="Calibri", size=11)

    ws.freeze_panes = ws.cell(row=fila_encabezado + 2, column=n_filas_agrupadoras + 1)
    ws.print_title_rows = f"1:{fila_encabezado}"
    for j, texto in enumerate(encabezados, start=1):
        ancho = max(13, min(40, len(str(texto)) + 6))
        ws.column_dimensions[get_column_letter(j)].width = ancho

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def fecha_desde_nombre_archivo(nombre: str) -> str | None:
    m = re.search(r"(\d{1,2})[-_]([A-ZÁÉÍÓÚa-záéíóú]+)[-_](\d{4})", nombre)
    if not m:
        return None
    dia, mes_txt, anio = m.groups()
    mes_txt = mes_txt.lower()
    meses_largos = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
                    "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
    mes_encontrado = next((m for m in meses_largos if m.startswith(mes_txt[:3])), None)
    if not mes_encontrado:
        return None
    return f"{int(dia)} de {mes_encontrado} de {anio}"


# ---------------------------------------------------------------------------
# Interfaz Streamlit
# ---------------------------------------------------------------------------
def main():
    st.set_page_config(page_title="Estado del Ejercicio — MAP / SICOP", layout="wide")
    st.title("Estado del Ejercicio — MAP / SICOP")
    st.caption(
        "Sube el MAP o SICOP y descarga el Estado del Ejercicio. "
    )

    with st.sidebar:
        st.header("1. Fuente de datos")
        fuente = st.radio("¿Qué vas a cargar?", ["MAP", "SICOP"], horizontal=True)
        archivo = st.file_uploader(f"Archivo crudo de {fuente} (.csv o .xlsx)", type=["csv", "xlsx", "xls"])
        st.divider()
        st.header("2. Periodo")
        hoy = date.today()
        mes_corte_idx = st.selectbox(
            "Corte de 'Al periodo' (acumulado enero → este mes)",
            options=list(range(12)),
            format_func=lambda i: NOMBRES_MES[i].capitalize(),
            index=min(hoy.month - 1, 11),
        )
        depurar_sicop = False
        if fuente == "SICOP":
            depurar_sicop = st.checkbox(
                "Excluir capítulo 1000, partida 39801 y CONTROL_OPERATIVO 60-69",
                value=False,
                help="Reglas confirmadas en tu procesador SICOP (nrbeca/nuevo). Revisa si aplican al reporte que quieres armar.",
            )

    if not archivo:
        st.info("Sube un archivo en la barra lateral para generar el reporte.")
        return

    try:
        df = cargar_crudo(archivo.getvalue(), archivo.name, fuente)
    except Exception as e:
        st.error(f"No se pudo leer el archivo: {e}")
        return

    df = enriquecer_con_catalogos(df)
    df, _ = agregar_periodos_y_disponible(df, fuente, mes_corte_idx)
    if fuente == "SICOP" and depurar_sicop:
        df = aplicar_depuracion_sicop(df)

    st.success(f"Archivo cargado: {archivo.name} — {len(df):,} filas")

    fecha_archivo = fecha_desde_nombre_archivo(archivo.name)
    titulo_default = (f"Estado del Ejercicio al {fecha_archivo}" if fecha_archivo
                       else f"Estado del Ejercicio al {hoy.day} de {NOMBRES_MES[hoy.month-1]} de {hoy.year}")

    # -----------------------------------------------------------------
    # Unidades a incluir en el cuadro. Por default salen TODAS las
    # unidades presentes en el archivo cargado (esto ya cubre todas las
    # OREF, 512, 513 y 120-811, porque son códigos de unidad como
    # cualquier otro). Si se desactiva "Todas las unidades", se puede
    # elegir cualquier combinación puntual (una sola UR, o varias juntas
    # como 512 + 513, o 120 + 811).
    # -----------------------------------------------------------------
    cat_ur_nombres = cargar_catalogo("unidades.csv", "codigo_ur", "nombre_ur")
    codigos_disponibles = (
        sorted(df["Unidad Responsable"].dropna().astype(str).unique(), key=lambda c: (len(c), c))
        if "Unidad Responsable" in df.columns else []
    )

    st.subheader("Unidades a incluir en el cuadro")
    todas_las_unidades = st.checkbox(
        "Todas las unidades (incluye todas las OREF, 512, 513 y 120-811)",
        value=True,
    )
    unidades_seleccionadas = codigos_disponibles
    if not todas_las_unidades:
        unidades_seleccionadas = st.multiselect(
            "Elige una o varias unidades (por ejemplo solo 512, solo 513, o 120 + 811 juntas)",
            options=codigos_disponibles,
            default=[],
            format_func=lambda c: etiqueta_con_nombre(c, cat_ur_nombres),
        )
        if not unidades_seleccionadas:
            st.info("Elige al menos una unidad, o activa 'Todas las unidades'.")
            return
        df = df[df["Unidad Responsable"].astype(str).isin(unidades_seleccionadas)]

    pivote, encabezados, grupos, filas, roles = construir_reporte_plantilla(df, fuente, mes_corte_idx)
    if pivote.empty:
        st.warning("La base cargada no trae las columnas necesarias (Unidad Responsable / Partida) para armar el reporte.")
        return

    st.dataframe(pivote, use_container_width=True, height=420)

    # Encabezados institucionales: se arman solos, ya no se capturan a mano.
    linea1 = "Unidad de Administración y Finanzas"
    titulo = titulo_default
    if not todas_las_unidades and len(unidades_seleccionadas) == 1:
        nombre_sel = solo_nombre(unidades_seleccionadas[0], cat_ur_nombres)
        linea2 = nombre_sel or "Dirección General de Programación, Presupuesto y Finanzas"
        subtitulo = f"Reporte {fuente} — UR {unidades_seleccionadas[0]}"
    else:
        linea2 = "Dirección General de Programación, Presupuesto y Finanzas"
        subtitulo = f"Reporte {fuente}"

    excel_bytes = exportar_excel_oref(pivote, fuente, linea1, linea2, titulo, subtitulo, filas, encabezados, grupos, roles)
    st.download_button(
        " Descargar Excel — Estado del Ejercicio",
        data=excel_bytes,
        file_name=f"Estado_del_Ejercicio_{fuente}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
