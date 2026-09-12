# -*- coding: utf-8 -*-
"""
nhv_processor.py
Módulo portátil para a lectura, análise e xeración da Ficha de Xustificación
das Normas de Habitabilidade de Vivendas de Galicia (NHV - Decreto 29/2010)
conservando exactamente os estilos, táboas e formatos orixinais de NHV.docx.
"""

import os
import sys
import io
import docx

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

def get_default_nhv_paths():
    """Localiza a plantilla orixinal NHV.docx de xeito portable."""
    cwd = os.getcwd()
    candidates = [
        os.path.join(cwd, "NHV.docx"),
        os.path.join(SCRIPT_DIR, "NHV.docx")
    ]
    for c in candidates:
        if os.path.exists(c):
            return c, cwd
    return os.path.join(cwd, "NHV.docx"), cwd

# Catálogo completo das 173 comprobacións do Decreto 29/2010 nas táboas 0 a 5
NHV_CATALOG = [
    {"id": "t0_r1", "table": 0, "row": 1, "section": "A.1.", "concepto": "A.1. CONDICIÓNS DE DESEÑO, CALIDADE E SOSTIBILIDADE", "title": "A.1.1. Condicións de vivenda exterior", "normativa": "A vivenda ten a consideración de VIVENDA EXTERIOR.", "default": "SI"},
    {"id": "t0_r2", "table": 0, "row": 2, "section": "A.1.", "concepto": "A.1. CONDICIÓNS DE DESEÑO, CALIDADE E SOSTIBILIDADE", "title": "Condicións definidas polo Plan. (1)", "normativa": "SI/NON", "default": "SI"},
    {"id": "t0_r3", "table": 0, "row": 3, "section": "A.1.", "concepto": "A.1. CONDICIÓNS DE DESEÑO, CALIDADE E SOSTIBILIDADE", "title": "A estancia maior -en todos os casos- e, cando a vivenda conte con máis dunha estancia, outra estancia ou a cociña -que non estea integrada na estancia maior- deberán ter iluminación e ventilación natural e relación co exterior a través de:", "normativa": "Rúas, prazas e espazos libres públicos definidos polo plan ou a normativa urbanística.", "default": "Non aplica"},
    {"id": "t0_r4", "table": 0, "row": 4, "section": "A.1.", "concepto": "A.1. CONDICIÓNS DE DESEÑO, CALIDADE E SOSTIBILIDADE", "title": "A estancia maior -en todos os casos- e, cando a vivenda conte con máis dunha estancia, outra estancia ou a cociña -que non estea integrada na estancia maior- deberán ter iluminación e ventilación natural e relación co exterior a través de:", "normativa": "Patios de cuarteirón ou espazos libres públicos ou privados onde sexa posible inscribir, fóra da proxección dos voos interiores, un círculo de Ø ≥ 0,7 H m. (2)", "default": "Non aplica"},
    {"id": "t0_r5", "table": 0, "row": 5, "section": "A.1.", "concepto": "A.1. CONDICIÓNS DE DESEÑO, CALIDADE E SOSTIBILIDADE", "title": "A estancia maior -en todos os casos- e, cando a vivenda conte con máis dunha estancia, outra estancia ou a cociña -que non estea integrada na estancia maior- deberán ter iluminación e ventilación natural e relación co exterior a través de:", "normativa": "Calquera espazo libre, público ou privado identificado como espazo exterior polo plan ou polo anexo de habitabilidade.", "default": "Non aplica"},
    {"id": "t0_r6", "table": 0, "row": 6, "section": "A.1.", "concepto": "A.1. CONDICIÓNS DE DESEÑO, CALIDADE E SOSTIBILIDADE", "title": "A estancia maior -en todos os casos- e, cando a vivenda conte con máis dunha estancia, outra estancia ou a cociña -que non estea integrada na estancia maior- deberán ter iluminación e ventilación natural e relación co exterior a través de:", "normativa": "En vivendas unifamiliares no espazo libre de parcela deberá poder inscribirse un círculo de Ø ≥ 0,5 H, mínimo 3 m.", "default": "Non aplica"},
    {"id": "t0_r7", "table": 0, "row": 7, "section": "A.1.", "concepto": "A.1. CONDICIÓNS DE DESEÑO, CALIDADE E SOSTIBILIDADE", "title": "", "normativa": "Toda peza vivideira ten iluminación natural e luz directa (7) desde o exterior a través dun dos espazos definidos no punto anterior, ou ben a través dos patios definidos no B.1.3, mediante unha xanela ou porta ubicada no plano da envolvente exterior.", "default": "SI"},
    {"id": "t1_r0", "table": 1, "row": 0, "section": "A.1.", "concepto": "A.1. CONDICIÓNS DE DESEÑO, CALIDADE E SOSTIBILIDADE", "title": "Superficie mínima de acristalamento para iluminación nas pezas vivideiras.", "normativa": "1/8 da superficie útil da peza.", "default": "Existente (Rehab.)"},
    {"id": "t1_r1", "table": 1, "row": 1, "section": "A.1.", "concepto": "A.1. CONDICIÓNS DE DESEÑO, CALIDADE E SOSTIBILIDADE", "title": "Altura máxima de antepeito en xanelas proxectadas para dar cumprimento ás condicións de habitabilidade, medida ata o pavimento rematado da peza.", "normativa": "1,10 m.", "default": "Existente (Rehab.)"},
    {"id": "t1_r2", "table": 1, "row": 2, "section": "A.1.", "concepto": "A.1. CONDICIÓNS DE DESEÑO, CALIDADE E SOSTIBILIDADE", "title": "Altura máxima do chan dos espazos exteriores aos que ventilen as estancias por enriba do pavimento rematado destas.", "normativa": "0,50 m.", "default": "Existente (Rehab.)"},
    {"id": "t1_r3", "table": 1, "row": 3, "section": "A.1.", "concepto": "A.1. CONDICIÓNS DE DESEÑO, CALIDADE E SOSTIBILIDADE", "title": "Protección de vistas establecida no Anexo de habitabilidade.", "normativa": "SI/NO", "default": "NON"},
    {"id": "t1_r4", "table": 1, "row": 4, "section": "A.1.", "concepto": "A.1. CONDICIÓNS DE DESEÑO, CALIDADE E SOSTIBILIDADE", "title": "Altura mínima da cara inferior do oco no que se aloxen as xanelas e calquera elemento transparente que se poida situar en dito oco e permita a visión do interior das pezas vivideiras.", "normativa": "1,80 m por enriba do chan do espazo exterior de uso público.", "default": "Non aplica"},
    {"id": "t1_r5", "table": 1, "row": 5, "section": "A.1.", "concepto": "A.1. CONDICIÓNS DE DESEÑO, CALIDADE E SOSTIBILIDADE", "title": "Fondo da franxa de terreo de uso privativo da vivenda entre a fachada na que se sitúa a xanela e o espazo público.", "normativa": "≥ 2 m.", "default": "≥ 2 m"},
    {"id": "t1_r6", "table": 1, "row": 6, "section": "A.1.", "concepto": "A.1. CONDICIÓNS DE DESEÑO, CALIDADE E SOSTIBILIDADE", "title": "Superficie mínima de iluminación.", "normativa": "1/6 da superficie útil da peza.", "default": "Non aplica"},
    {"id": "t1_r7", "table": 1, "row": 7, "section": "A.1.", "concepto": "A.1. CONDICIÓNS DE DESEÑO, CALIDADE E SOSTIBILIDADE", "title": "Profundidade máxima.", "normativa": "3 m.", "default": "Non aplica"},
    {"id": "t1_r8", "table": 1, "row": 8, "section": "A.1.", "concepto": "A.1. CONDICIÓNS DE DESEÑO, CALIDADE E SOSTIBILIDADE", "title": "Lonxitude ≥ profundidade.", "normativa": "SI", "default": "Non aplica"},
    {"id": "t1_r9", "table": 1, "row": 9, "section": "A.1.", "concepto": "A.1. CONDICIÓNS DE DESEÑO, CALIDADE E SOSTIBILIDADE", "title": "Superficie mínima de iluminación.", "normativa": "1/6 da superficie útil da peza.", "default": "Non aplica"},
    {"id": "t1_r10", "table": 1, "row": 10, "section": "A.1.", "concepto": "A.1. CONDICIÓNS DE DESEÑO, CALIDADE E SOSTIBILIDADE", "title": "Mantense a continuidade da envolvente principal.", "normativa": "SI", "default": "Non aplica"},
    {"id": "t1_r11", "table": 1, "row": 11, "section": "A.1.", "concepto": "A.1. CONDICIÓNS DE DESEÑO, CALIDADE E SOSTIBILIDADE", "title": "P ≤ 7.50 m", "normativa": "1/8 da superficie útil da peza.", "default": "Existente (Rehab.)"},
    {"id": "t1_r12", "table": 1, "row": 12, "section": "A.1.", "concepto": "A.1. CONDICIÓNS DE DESEÑO, CALIDADE E SOSTIBILIDADE", "title": "7,50 m < P < 2,2 A (3)", "normativa": "1/6 da superficie útil da peza.", "default": "Existente (Rehab.)"},
    {"id": "t1_r13", "table": 1, "row": 13, "section": "A.1.", "concepto": "A.1. CONDICIÓNS DE DESEÑO, CALIDADE E SOSTIBILIDADE", "title": "Superficie mínima para iluminación.", "normativa": "1/8 da superficie útil da peza.", "default": "Existente (Rehab.)"},
    {"id": "t1_r14", "table": 1, "row": 14, "section": "A.1.", "concepto": "A.1. CONDICIÓNS DE DESEÑO, CALIDADE E SOSTIBILIDADE", "title": "Altura desde a parte inferior da xanela ata o pavimento rematado da estancia.", "normativa": "≤ 1,20 m.", "default": "Existente (Rehab.)"},
    {"id": "t1_r15", "table": 1, "row": 15, "section": "A.1.", "concepto": "A.1. CONDICIÓNS DE DESEÑO, CALIDADE E SOSTIBILIDADE", "title": "Altura desde a parte superior da xanela ata o pavimento rematado da estancia.", "normativa": "≥ 2,00 m.", "default": "Existente (Rehab.)"},
    {"id": "t1_r16", "table": 1, "row": 16, "section": "A.1.", "concepto": "A.1. CONDICIÓNS DE DESEÑO, CALIDADE E SOSTIBILIDADE", "title": "Superficie mínima real de ventilación nas pezas vivideiras.", "normativa": "1/3 da superficie mínima de iluminación.", "default": "Existente (Rehab.)"},
    {"id": "t1_r17", "table": 1, "row": 17, "section": "A.1.", "concepto": "A.1. CONDICIÓNS DE DESEÑO, CALIDADE E SOSTIBILIDADE", "title": "Se manteñen os ocos de iluminación e ventilación existentes en obras de remodelación de vivendas e obras de adecuación funcional de edificios.", "normativa": "SI/NON", "default": "SI"},
    {"id": "t1_r18", "table": 1, "row": 18, "section": "A.1.", "concepto": "A.1. CONDICIÓNS DE DESEÑO, CALIDADE E SOSTIBILIDADE", "title": "As determinacións da normativa urbanística ou de protección do patrimonio non permiten o seu cumprimento.", "normativa": "SI/NON", "default": "NON"},
    {"id": "t1_r19", "table": 1, "row": 19, "section": "A.1.", "concepto": "A.1. CONDICIÓNS DE DESEÑO, CALIDADE E SOSTIBILIDADE", "title": "Só é exixible o cumprimento das condición de protección de vistas á estancia maior, en todos os casos, e a outra estancia, no caso de que a vivenda conte con máis dunha.", "normativa": "SI/NON", "default": "SI"},
    {"id": "t1_r20", "table": 1, "row": 20, "section": "A.2.", "concepto": "A.2. CONDICIÓNS FUNCIONAIS", "title": "A vivenda ten acceso desde un espazo público ou desde un espazo común do edificio ou urbanización con comunicación directa co espazo público:", "normativa": "Directo.", "default": "Non aplica"},
    {"id": "t1_r21", "table": 1, "row": 21, "section": "A.2.", "concepto": "A.2. CONDICIÓNS FUNCIONAIS", "title": "A vivenda ten acceso desde un espazo público ou desde un espazo común do edificio ou urbanización con comunicación directa co espazo público:", "normativa": "A través dun anexo vinculado a ela.", "default": "Non aplica"},
    {"id": "t1_r22", "table": 1, "row": 22, "section": "A.2.", "concepto": "A.2. CONDICIÓNS FUNCIONAIS", "title": "A vivenda ten acceso desde un espazo público ou desde un espazo común do edificio ou urbanización con comunicación directa co espazo público:", "normativa": "A través dunha parcela da súa propiedade.", "default": "SI"},
    {"id": "t1_r23", "table": 1, "row": 23, "section": "A.2.", "concepto": "A.2. CONDICIÓNS FUNCIONAIS", "title": "A vivenda ten acceso desde un espazo público ou desde un espazo común do edificio ou urbanización con comunicación directa co espazo público:", "normativa": "A través dunha parcela sobre a que se ten dereito de paso.", "default": "Non aplica"},
    {"id": "t2_r0", "table": 2, "row": 0, "section": "A.2.", "concepto": "A.2. CONDICIÓNS FUNCIONAIS", "title": "", "normativa": "A vivenda é paso obrigado para acceder a calquera local ou parcela que non sexa de uso exclusivo da mesma.", "default": "NON"},
    {"id": "t2_r1", "table": 2, "row": 1, "section": "A.2.", "concepto": "A.2. CONDICIÓNS FUNCIONAIS", "title": "", "normativa": "As dependencias da vivenda comunícanse entre si a través de espazos pechados de uso exclusivo dos seus moradores.", "default": "SI"},
    {"id": "t2_r2", "table": 2, "row": 2, "section": "A.2.", "concepto": "A.2. CONDICIÓNS FUNCIONAIS", "title": "A.2.2. Composición e comparti-mentación", "normativa": "Paso ás pezas vivideiras desde o acceso á vivenda a través da estancia maior e os espazos de comunicación.", "default": "SI"},
    {"id": "t2_r3", "table": 2, "row": 3, "section": "A.2.", "concepto": "A.2. CONDICIÓNS FUNCIONAIS", "title": "Paso obrigado a outras estancias ou cociña a través da estancia maior (salvo que a cociña estea integrada na estancia maior e esta non sexa de paso obrigado para ningunha outra estancia).", "normativa": "Aumento da superficie da estancia maior de 2 m2.", "default": "SI"},
    {"id": "t2_r4", "table": 2, "row": 4, "section": "A.2.", "concepto": "A.2. CONDICIÓNS FUNCIONAIS", "title": "O acceso ao cuarto de baño obrigatorio efectuarase a través dos espazos de comunicación da vivenda.", "normativa": "SI/NON", "default": "SI"},
    {"id": "t2_r5", "table": 2, "row": 5, "section": "A.2.", "concepto": "A.2. CONDICIÓNS FUNCIONAIS", "title": "Se a vivenda tamén inclúe un aseo -que conte con ducha ou bañeira- o acceso a un deles deberá efectuarase a través dos espazos de comunicación, podendo accederse ó outro tamén a través dunha estancia. Neste caso, se o acceso ao baño se realiza a través dunha estancia, o aseo deberá cumprir coas súas determinacións.", "normativa": "SI", "default": "Baño dende zona común"},
    {"id": "t2_r6", "table": 2, "row": 6, "section": "A.2.", "concepto": "A.2. CONDICIÓNS FUNCIONAIS", "title": "En vivendas de dúas estancias, o acceso ao cuarto de baño obrigatorio poderá efectuarse desde a segunda estancia, aínda que a vivenda non conte cun aseo.", "normativa": "SI", "default": "Non aplica"},
    {"id": "t2_r7", "table": 2, "row": 7, "section": "A.2.", "concepto": "A.2. CONDICIÓNS FUNCIONAIS", "title": "En vivendas dunha estancia, o acceso ao cuarto de baño obrigatorio poderá realizarse desde espazos comúns ou a través da estancia, sempre que o baño estea dividido en dúas pezas.", "normativa": "SI", "default": "Non aplica"},
    {"id": "t2_r8", "table": 2, "row": 8, "section": "A.2.", "concepto": "A.2. CONDICIÓNS FUNCIONAIS", "title": "A.2.3. Programa mínimo", "normativa": "Unha estancia, unha cociña, un cuarto de baño, un lavadoiro, un tendal e un espazo de almacenamento.", "default": "SI"},
    {"id": "t2_r9", "table": 2, "row": 9, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Altura libre mínima entre pavimento e teito acabados.", "normativa": "2,50 m.", "default": "≥ 2,50 m"},
    {"id": "t2_r10", "table": 2, "row": 10, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "A altura anterior pódese diminuír ata 2,20 m.", "normativa": "Ata o 30% da superficie útil.", "default": "≤ 30 %"},
    {"id": "t2_r11", "table": 2, "row": 11, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Altura libre mínima entre pavimento e teito acabados en vestíbulos, corredores, escaleiras, cuartos de baño, aseos, lavadoiros, tendais e garaxes de vivendas unifamiliares.", "normativa": "2,20 m.", "default": "≥ 2,20 m"},
    {"id": "t2_r12", "table": 2, "row": 12, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "*REHABIILITACION de edificios ou vivendas, agás que se modifique a posición dos forxados existentes.", "normativa": "Pódense manter as alturas existentes.", "default": "h existente cumpre"},
    {"id": "t2_r13", "table": 2, "row": 13, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "*REHABIILITACION: Altura libre mínima admisible entre pavimento e teito acabados no caso de cambio de uso a vivenda de locais ou espazos que non tiñan o devandito uso.", "normativa": "2,40 m.", "default": "h existente cumpre"},
    {"id": "t2_r14", "table": 2, "row": 14, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "O volume mínimo da peza é igual á súa superficie útil mínima multiplicada pola súa altura exixible.", "normativa": "SI", "default": ""},
    {"id": "t2_r15", "table": 2, "row": 15, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "% da superficie mínima esixible á peza que ten unha altura ≥ 2,50 m (estancias/cociñas) ou 2,20 m (aseos/baños…).", "normativa": "≥ 70%.", "default": "Non aplica"},
    {"id": "t2_r16", "table": 2, "row": 16, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Altura mínima de corredores e vestíbulos abufardados que sirvan de acceso ás pezas.", "normativa": "2,20 m.", "default": "Non aplica"},
    {"id": "t2_r17", "table": 2, "row": 17, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Altura mínima libre do espazo ocupado polo Cadrado Base.", "normativa": "1,80 m.", "default": "Non aplica"},
    {"id": "t2_r19", "table": 2, "row": 19, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "A.3.2. DIMENSIÓNS SUPERFICIAIS E LINEAIS", "normativa": "A.3.2. DIMENSIÓNS SUPERFICIAIS E LINEAIS", "default": "Estancia maior, E1"},
    {"id": "t2_r20", "table": 2, "row": 20, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie útil mínima da estancia E1 para nº estancias =1", "normativa": "25,00 m2.", "default": "Non aplica"},
    {"id": "t2_r21", "table": 2, "row": 21, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie útil mínima da estancia E1 para nº estancias =2", "normativa": "16,00 m2.", "default": "Non aplica"},
    {"id": "t2_r22", "table": 2, "row": 22, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie útil mínima da estancia E1 para nº estancias =3", "normativa": "18,00 m2.", "default": "Non aplica"},
    {"id": "t2_r23", "table": 2, "row": 23, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie útil mínima da estancia E1 para nº estancias =4", "normativa": "20,00 m2.", "default": "Non aplica"},
    {"id": "t2_r24", "table": 2, "row": 24, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie útil mínima da estancia E1 para nº estancias =5", "normativa": "22,00 m2.", "default": "Non aplica"},
    {"id": "t2_r25", "table": 2, "row": 25, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie útil mínima da estancia E1 para nº estancias >5", "normativa": "25,00 m2.", "default": "61,35 m2"},
    {"id": "t2_r26", "table": 2, "row": 26, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Redución da superficie de E1 por aumentar a superficie da cociña en 4 m2 ou máis.", "normativa": "≤ 4 m2.", "default": "Non aplica"},
    {"id": "t2_r27", "table": 2, "row": 27, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Cadrado Base inscritible na súa planta. (4)", "normativa": "3,30 m de lado.", "default": "3,30 x 3,30 m"},
    {"id": "t3_r0", "table": 3, "row": 0, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie total de elementos puntuais admisibles que non sobresaian máis de 0,30 m (nun ou máis lados do cadrado)", "normativa": "0,15 m2.", "default": "0,00 m2"},
    {"id": "t3_r1", "table": 3, "row": 1, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Ancho mínimo entre paramentos enfrontados.", "normativa": "2,70 m.", "default": "4,56 m"},
    {"id": "t3_r2", "table": 3, "row": 2, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "A.3.2. DIMENSIÓNS SUPERFICIAIS E LINEAIS", "normativa": "A.3.2. DIMENSIÓNS SUPERFICIAIS E LINEAIS", "default": "Estancia E2"},
    {"id": "t3_r3", "table": 3, "row": 3, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie útil mínima da estancia E2 para calquera nº de estancias.", "normativa": "≥ 12,00 m2.", "default": "32,25 m2"},
    {"id": "t3_r4", "table": 3, "row": 4, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Cadrado Base inscritible na súa planta. (4)", "normativa": "2,60 m de lado.", "default": "2,60x2,60 m"},
    {"id": "t3_r5", "table": 3, "row": 5, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie total de elementos puntuais admisibles que non sobresaian máis de 0,30 m (nun só lado do cadrado).", "normativa": "0,15 m2.", "default": "0,00 m2"},
    {"id": "t3_r6", "table": 3, "row": 6, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Ancho mínimo entre paramentos enfrontados.", "normativa": "2,60 m.", "default": "5,27 m"},
    {"id": "t3_r7", "table": 3, "row": 7, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "% de superficie útil de espazos de acceso á estancia, con anchos inferiores a 2,60 m entre paramentos, pero que computan a efectos de superficie mínima porque serven como acceso directo ao almacenamento ou baños/aseos complementarios da mesma.", "normativa": "≤ 10% da superficie útil da estancia.", "default": "0%"},
    {"id": "t3_r8", "table": 3, "row": 8, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "A.3.2. DIMENSIÓNS SUPERFICIAIS E LINEAIS", "normativa": "A.3.2. DIMENSIÓNS SUPERFICIAIS E LINEAIS", "default": "Estancia E3"},
    {"id": "t3_r9", "table": 3, "row": 9, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie útil mínima da estancia E3 para calquera nº de estancias.", "normativa": "8,00 m2.", "default": "12,90 m2"},
    {"id": "t3_r10", "table": 3, "row": 10, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Cadrado Base inscritible na súa planta. (4)", "normativa": "2,20 m de lado.", "default": "2,20x2,20 m"},
    {"id": "t3_r11", "table": 3, "row": 11, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie total de elementos puntuais admisibles que non sobresaian máis de 0,30 m (nun só lado do cadrado).", "normativa": "0,15 m2.", "default": "0,00 m2"},
    {"id": "t3_r12", "table": 3, "row": 12, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Ancho mínimo entre paramentos enfrontados.", "normativa": "2,00 m.", "default": "2,79 m"},
    {"id": "t3_r13", "table": 3, "row": 13, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "% de superficie útil de espazos de acceso á estancia, con anchos inferiores a 2,00 m entre paramentos, pero que computan a efectos de superficie mínima porque serven como acceso directo ao almacenamento ou baños/aseos complementarios da mesma.", "normativa": "≤ 10% da superficie útil da estancia.", "default": "0%"},
    {"id": "t3_r14", "table": 3, "row": 14, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "A.3.2. DIMENSIÓNS SUPERFICIAIS E LINEAIS", "normativa": "A.3.2. DIMENSIÓNS SUPERFICIAIS E LINEAIS", "default": "Estancia E4"},
    {"id": "t3_r15", "table": 3, "row": 15, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie útil mínima da estancia E4 para calquera nº de estancias.", "normativa": "8,00 m2.", "default": "12,20 m2"},
    {"id": "t3_r16", "table": 3, "row": 16, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Cadrado Base inscritible na súa planta. (4)", "normativa": "2,20 m de lado.", "default": "2,20x2,20 m"},
    {"id": "t3_r17", "table": 3, "row": 17, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie total de elementos puntuais admisibles que non sobresaian máis de 0,30 m (nun só lado do cadrado).", "normativa": "0,15 m2.", "default": "0,00 m2"},
    {"id": "t3_r18", "table": 3, "row": 18, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Ancho mínimo entre paramentos enfrontados.", "normativa": "2,00 m.", "default": "2,81 m"},
    {"id": "t3_r19", "table": 3, "row": 19, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "% de superficie útil de espazos de acceso á estancia, con anchos inferiores a 2,00 m entre paramentos, pero que computan a efectos de superficie mínima porque serven como acceso directo ao almacenamento persoal ou baños/aseos complementarios da mesma.", "normativa": "≤ 10% da superficie útil da estancia.", "default": "0%"},
    {"id": "t3_r20", "table": 3, "row": 20, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "A.3.2. DIMENSIÓNS SUPERFICIAIS E LINEAIS", "normativa": "A.3.2. DIMENSIÓNS SUPERFICIAIS E LINEAIS", "default": "Estancia E5"},
    {"id": "t3_r21", "table": 3, "row": 21, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie útil mínima da estancia E5 para nº estancias =5", "normativa": "6,00 m2.", "default": "Non aplica"},
    {"id": "t3_r22", "table": 3, "row": 22, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie útil mínima da estancia E5 para nº estancias >5", "normativa": "8,00 m2.", "default": "10,77 m2"},
    {"id": "t3_r23", "table": 3, "row": 23, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Cadrado Base inscritible na súa planta. (4)", "normativa": "2,20 m de lado.", "default": "2,20x2,20 m"},
    {"id": "t3_r24", "table": 3, "row": 24, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie total de elementos puntuais admisibles que non sobresaian máis de 0,30 m (nun só lado do cadrado).", "normativa": "0,15 m2.", "default": "0,00 m2"},
    {"id": "t3_r25", "table": 3, "row": 25, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Ancho mínimo entre paramentos enfrontados.", "normativa": "2,00 m.", "default": "2,79 m"},
    {"id": "t3_r26", "table": 3, "row": 26, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "% de superficie útil de espazos de acceso á estancia, con anchos inferiores a 2,00 m entre paramentos, pero que computan a efectos de superficie mínima porque serven como acceso directo ao almacenamento ou baños/aseos complementarios da mesma.", "normativa": "≤ 10% da superficie útil da estancia.", "default": "0%"},
    {"id": "t3_r27", "table": 3, "row": 27, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "A.3.2. DIMENSIÓNS SUPERFICIAIS E LINEAIS", "normativa": "A.3.2. DIMENSIÓNS SUPERFICIAIS E LINEAIS", "default": "Estancia E6"},
    {"id": "t3_r28", "table": 3, "row": 28, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie útil mínima de estancia En para nº estancias >5", "normativa": "6,00 m2.", "default": "9,86 m2"},
    {"id": "t3_r29", "table": 3, "row": 29, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Cadrado Base inscritible na súa planta. (4)", "normativa": "2,20 m de lado.", "default": "2,20x2,20 m"},
    {"id": "t3_r30", "table": 3, "row": 30, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie total de elementos puntuais admisibles que non sobresaian máis de 0,30 m (nun só lado do cadrado).", "normativa": "0,15 m2.", "default": "0,00 m2"},
    {"id": "t3_r31", "table": 3, "row": 31, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Ancho mínimo entre paramentos enfrontados", "normativa": "2,00 m.", "default": "2,79 m"},
    {"id": "t3_r32", "table": 3, "row": 32, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "% de superficie útil de espazos de acceso á estancia, con anchos inferiores a 2,00 m entre paramentos, pero que computan a efectos de superficie mínima porque serven como acceso directo ao almacenamento ou baños/aseos complementarios da mesma.", "normativa": "≤ 10% da superficie útil da estancia.", "default": "0%"},
    {"id": "t4_r0", "table": 4, "row": 0, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "A.3.2. DIMENSIÓNS SUPERFICIAIS E LINEAIS", "normativa": "A.3.2. DIMENSIÓNS SUPERFICIAIS E LINEAIS", "default": "Outras estancias"},
    {"id": "t4_r1", "table": 4, "row": 1, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "A superficie útil computable a efectos de habitabilidade do conxunto das estancias da vivenda supera os 100 m2.", "normativa": "SI/NON", "default": "SI"},
    {"id": "t4_r2", "table": 4, "row": 2, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Se NON supera os 100 m2, as pezas distintas dos servizos de superficie > 3 m2 deben cumprir as condicións establecidas para as estancias.", "normativa": "SI", "default": "Non aplica"},
    {"id": "t4_r4", "table": 4, "row": 4, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "A.3.2. DIMENSIÓNS SUPERFICIAIS E LINEAIS", "normativa": "A.3.2. DIMENSIÓNS SUPERFICIAIS E LINEAIS", "default": "Cociña"},
    {"id": "t4_r5", "table": 4, "row": 5, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie útil mínima da cociña para nº estancias =1", "normativa": "5,00 m2.", "default": "Non aplica"},
    {"id": "t4_r6", "table": 4, "row": 6, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie útil mínima da cociña para nº estancias =2", "normativa": "7,00 m2.", "default": "Non aplica"},
    {"id": "t4_r7", "table": 4, "row": 7, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie útil mínima da cociña para nº estancias =3", "normativa": "7,00 m2.", "default": "Non aplica"},
    {"id": "t4_r8", "table": 4, "row": 8, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie útil mínima da cociña para nº estancias =4", "normativa": "9,00 m2.", "default": "Non aplica"},
    {"id": "t4_r9", "table": 4, "row": 9, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie útil mínima da cociña para nº estancias =5", "normativa": "9,00 m2.", "default": "Non aplica"},
    {"id": "t4_r10", "table": 4, "row": 10, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie útil mínima da cociña para nº estancias >5", "normativa": "10,00 m2.", "default": ""},
    {"id": "t4_r11", "table": 4, "row": 11, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Cando a cociña se integra nun único espazo coa estancia maior, esta última conserva a súa superficie mínima, cada zona cumpre a súa distancia mínima entre paramentos e a superficie mínima das dúas pezas integradas será:", "normativa": "A suma das superficies mínimas establecidas para cada unha das pezas.", "default": "61,35 m2 (>25+10+2)"},
    {"id": "t4_r12", "table": 4, "row": 12, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Cociña integrada na estancia maior: superficie vertical aberta de relación entre estes espazos nun único oco.", "normativa": "≥ 3,5 m2.", "default": "7,89 m2"},
    {"id": "t4_r13", "table": 4, "row": 13, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Ancho mínimo entre paramentos enfrontados libre de obstáculos.", "normativa": "1,80 m.", "default": "1,91 m"},
    {"id": "t4_r14", "table": 4, "row": 14, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Lonxitude mínima fronte dedicado a mesado (sen contar o espazo destinado ao frigorífico).", "normativa": "2,40 m (superficie < 7 m2).", "default": "Non aplica"},
    {"id": "t4_r15", "table": 4, "row": 15, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Lonxitude mínima fronte dedicado a mesado (sen contar o espazo destinado ao frigorífico).", "normativa": "3,00 m (superficie ≥ 7 m2).", "default": "3,77 m"},
    {"id": "t4_r16", "table": 4, "row": 16, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Paso libre mínimo entre mesados e aparellos enfrontados.", "normativa": "0,90 m.", "default": "0,90 m"},
    {"id": "t4_r17", "table": 4, "row": 17, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "No caso de aumento da superficie da cociña de 4 m2, deberá poder inscribirse nela un Cadrado Base (4), non invadido polo mesado, de lado.", "normativa": "≥ 2,20 m.", "default": "2,20x2,20 m"},
    {"id": "t4_r18", "table": 4, "row": 18, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie total de elementos puntuais admisibles que non sobresaian máis de 0,30 m (nun só lado do cadrado).", "normativa": "0,15 m2.", "default": "0,00 m2"},
    {"id": "t4_r19", "table": 4, "row": 19, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie de espazos da cociña situados na súa entrada, con distancias entre paramentos enfrontados inferiores a 1,80 m, pero que computan a efectos de superficie mínima porque sirve de acceso a outros usos complementarios da mesma.", "normativa": "≤ 10% da superficie útil da cociña.", "default": "0%"},
    {"id": "t4_r20", "table": 4, "row": 20, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "A.3.2. DIMENSIÓNS SUPERFICIAIS E LINEAIS", "normativa": "A.3.2. DIMENSIÓNS SUPERFICIAIS E LINEAIS", "default": "Cuarto de baño"},
    {"id": "t4_r21", "table": 4, "row": 21, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie útil mínima.", "normativa": "5,00 m2.", "default": "5,49 m2"},
    {"id": "t4_r22", "table": 4, "row": 22, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Ancho mínimo entre paramentos enfrontados.", "normativa": "1,60 m.", "default": "1,97 m"},
    {"id": "t4_r23", "table": 4, "row": 23, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie total de elementos puntuais admisibles que non sobresaian máis de 0,20 m (nun só lado da peza).", "normativa": "0,15 m2.", "default": "0,00 m2"},
    {"id": "t4_r24", "table": 4, "row": 24, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "A.3.2. DIMENSIÓNS SUPERFICIAIS E LINEAIS", "normativa": "Disposición dos aparellos sanitarios que permita convertelo en baño de uso accesible no futuro segundo a normativa de accesibilidade vixente.", "default": "SI"},
    {"id": "t4_r25", "table": 4, "row": 25, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "O cuarto de baño poderá dividirse en dúas pezas comunicadas entre si. Na primeira situarase o lavabo e na segunda emprazaranse o resto dos aparellos sanitarios; o conxunto de ambas pezas deberá ser convertible en baño accesible no futuro. A suma das súas superficies deberá ser ≥ 5,00 m2 e a segunda peza deberá cumprir coa dimensión mínima entre paramentos enfrontados de 1,60 m.", "normativa": "SI", "default": "Non aplica"},
    {"id": "t4_r26", "table": 4, "row": 26, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "A.3.2. DIMENSIÓNS SUPERFICIAIS E LINEAIS", "normativa": "A.3.2. DIMENSIÓNS SUPERFICIAIS E LINEAIS", "default": "Cuarto de aseo"},
    {"id": "t4_r27", "table": 4, "row": 27, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie útil mínima do aseo cando sexa de obriga.", "normativa": "1,50 m2.", "default": "3,55 m2"},
    {"id": "t4_r28", "table": 4, "row": 28, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Ancho mínimo entre paramentos enfrontados.", "normativa": "1,20 m.", "default": "1,80 m"},
    {"id": "t4_r29", "table": 4, "row": 29, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie total de elementos puntuais admisibles que non sobresaian máis de 0,10 m (nun só lado da peza).", "normativa": "0,05 m2.", "default": "0,00 m2"},
    {"id": "t4_r30", "table": 4, "row": 30, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "A.3.2. DIMENSIÓNS SUPERFICIAIS E LINEAIS", "normativa": "A.3.2. DIMENSIÓNS SUPERFICIAIS E LINEAIS", "default": "Lavadoiro"},
    {"id": "t4_r31", "table": 4, "row": 31, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie útil mínima.", "normativa": "1,50 m2.", "default": "m2"},
    {"id": "t4_r32", "table": 4, "row": 32, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Ancho mínimo entre paramentos enfrontados.", "normativa": "1,20 m.", "default": "m"},
    {"id": "t4_r33", "table": 4, "row": 33, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "En todos os casos.", "normativa": "Desde os espazos de comunicación, as cociñas ou os cuartos de baño e aseo.", "default": "Non aplica"},
    {"id": "t4_r34", "table": 4, "row": 34, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Se a vivenda ten unha única estancia coa cociña integrada.", "normativa": "Tamén desde o devandito espazo.", "default": "Non aplica"},
    {"id": "t4_r35", "table": 4, "row": 35, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Nas vivendas unifamiliares, o acceso ao lavadoiro-tendal, ou a calquera deles se estivesen separados, poderá realizarse desde o garaxe -se este comunica co interior da vivenda-, desde outros espazos do interior do edificio -sempre que non sexan estancias- ou desde espazos cubertos da edificación principal.", "normativa": "SI", "default": "Non aplica"},
    {"id": "t4_r36", "table": 4, "row": 36, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Non é preciso reservar este espazo destinado a lavadoiro.", "normativa": "SI/NON", "default": "Non aplica"},
    {"id": "t4_r37", "table": 4, "row": 37, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie útil mínima cando o lavadoiro está integrado co tendal formando un espazo único.", "normativa": "≥ 3,00 m2.", "default": "Non aplica"},
    {"id": "t4_r38", "table": 4, "row": 38, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "A.3.2. DIMENSIÓNS SUPERFICIAIS E LINEAIS", "normativa": "A.3.2. DIMENSIÓNS SUPERFICIAIS E LINEAIS", "default": "Tendal"},
    {"id": "t4_r39", "table": 4, "row": 39, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie útil mínima.", "normativa": "1,50 m2.", "default": "Non aplica"},
    {"id": "t4_r40", "table": 4, "row": 40, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Ancho mínimo entre paramentos enfrontados.", "normativa": "1,20 m.", "default": "Non aplica"},
    {"id": "t4_r41", "table": 4, "row": 41, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Está cuberto e protexido de vistas desde o espazo público.", "normativa": "SI", "default": "Non aplica"},
    {"id": "t4_r42", "table": 4, "row": 42, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Interfire na ventilación/iluminación das pezas vivideiras.", "normativa": "NON", "default": "Non aplica"},
    {"id": "t4_r43", "table": 4, "row": 43, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Directa desde espazo exterior ou patio.", "normativa": "SI", "default": "Non aplica"},
    {"id": "t4_r44", "table": 4, "row": 44, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Situación fóra da envolvente térmica do edificio.", "normativa": "SI", "default": "Non aplica"},
    {"id": "t4_r45", "table": 4, "row": 45, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Ventilación permanente.", "normativa": "SI", "default": "Non aplica"},
    {"id": "t4_r46", "table": 4, "row": 46, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie mínima de ventilación.", "normativa": "1,5 m2.", "default": "Non aplica"},
    {"id": "t4_r47", "table": 4, "row": 47, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Se ventila a través de patio interior: Superficie mínima do conduto de entrada de aire desde o exterior na parte inferior do patio.", "normativa": "0,20 m2.", "default": "Non aplica"},
    {"id": "t4_r48", "table": 4, "row": 48, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Conta con calefacción.", "normativa": "SI", "default": "Non aplica"},
    {"id": "t4_r49", "table": 4, "row": 49, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Paredes revestidas de material impermeable á auga en toda a súa altura.", "normativa": "SI", "default": "Non aplica"},
    {"id": "t4_r50", "table": 4, "row": 50, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Condicións de ventilación: as establecidas no DB HS3 do CTE para aseos e cuartos de baño.", "normativa": "SI", "default": "Non aplica"},
    {"id": "t4_r51", "table": 4, "row": 51, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Nas vivendas unifamiliares con parcela propia, o espazo para secado da roupa poderá dispoñerse na parcela debendo quedar garantida a protección de vistas desde a rúa ao espazo público, a ventilación e a protección fronte a auga de choiva.", "normativa": "SI/NON", "default": "Non aplica"},
    {"id": "t4_r52", "table": 4, "row": 52, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Non é preciso reservar este espazo destinado a tendal.", "normativa": "SI/NON", "default": "SI"},
    {"id": "t4_r53", "table": 4, "row": 53, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "A.3.2. DIMENSIÓNS SUPERFICIAIS E LINEAIS", "normativa": "A.3.2. DIMENSIÓNS SUPERFICIAIS E LINEAIS", "default": "Espazos de almacenamento"},
    {"id": "t4_r54", "table": 4, "row": 54, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie útil mínima para nº estancias =1", "normativa": "1,00 m2.", "default": "Non aplica"},
    {"id": "t4_r55", "table": 4, "row": 55, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie útil mínima para nº estancias =2", "normativa": "2,00 m2.", "default": "Non aplica"},
    {"id": "t4_r56", "table": 4, "row": 56, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie útil mínima para nº estancias =3", "normativa": "3,00 m2.", "default": "Non aplica"},
    {"id": "t4_r57", "table": 4, "row": 57, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie útil mínima para nº estancias =4", "normativa": "4,00 m2.", "default": "Non aplica"},
    {"id": "t4_r58", "table": 4, "row": 58, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie útil mínima para nº estancias =5", "normativa": "5,00 m2.", "default": "Non aplica"},
    {"id": "t4_r59", "table": 4, "row": 59, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie útil mínima para nº estancias >5", "normativa": "6,00 m2.", "default": "6,03 m2"},
    {"id": "t4_r60", "table": 4, "row": 60, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Altura mínima do espazo de almacenamento.", "normativa": ".", "default": "2,20 m"},
    {"id": "t4_r61", "table": 4, "row": 61, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Fondo do espazo de almacenamento.", "normativa": "0,60 m ≤ F ≤ 0,75 m.", "default": "0,60 m"},
    {"id": "t4_r62", "table": 4, "row": 62, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Superficie útil mínima na que pode fraccionarse.", "normativa": "0,50 m2.", "default": "0,67 m2"},
    {"id": "t4_r63", "table": 4, "row": 63, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "En todos os casos.", "normativa": "Espazos de comunicación da vivenda (corredores e vestíbulos) ou en calquera estancia distinta da maior.", "default": "Dormitorios"},
    {"id": "t4_r64", "table": 4, "row": 64, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Nas vivendas de 1 ou 2 estancias.", "normativa": "Tamén na estancia maior.", "default": "Dormitorios"},
    {"id": "t4_r66", "table": 4, "row": 66, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "A.3.2. DIMENSIÓNS SUPERFICIAIS E LINEAIS", "normativa": "A.3.2. DIMENSIÓNS SUPERFICIAIS E LINEAIS", "default": "Espazo de acceso interior da vivenda"},
    {"id": "t4_r67", "table": 4, "row": 67, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Lado do cadrado a inscribir en contacto coa porta de entrada e libre de obstáculos. (6)", "normativa": ".", "default": "1,50x1,50 m"},
    {"id": "t4_r68", "table": 4, "row": 68, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "A.3.2. DIMENSIÓNS SUPERFICIAIS E LINEAIS", "normativa": "A.3.2. DIMENSIÓNS SUPERFICIAIS E LINEAIS", "default": "Corredores e zonas de acceso interiores as pezas"},
    {"id": "t4_r69", "table": 4, "row": 69, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Ancho libre mínimo entre paramentos.", "normativa": "1,00 m.", "default": "1,60 m"},
    {"id": "t4_r70", "table": 4, "row": 70, "section": "A.3.", "concepto": "A.3. CONDICIÓNS ESPACIAIS E DIMENSIONAIS DAS ESTANCIAS, SERVIZOS E ESPAZOS DE COMUNICACIÓN", "title": "Ancho libre mínimo con estreitamentos puntuais.", "normativa": "0,90 m.", "default": "1,60 m"},
    {"id": "t5_r1", "table": 5, "row": 1, "section": "A", "concepto": "", "title": "Ancho libre mínimo das portas de paso.", "normativa": ".", "default": "0,80 m"},
    {"id": "t5_r2", "table": 5, "row": 2, "section": "A", "concepto": "", "title": "Altura libre mínima das portas de paso.", "normativa": ".", "default": "2,03 m"},
    {"id": "t5_r3", "table": 5, "row": 3, "section": "A.4.", "concepto": "A.4. CONDICIÓNS DOTACIONAIS DAS VIVENDAS", "title": "A.4.1. DOTACIÓN MÍNIMA NA VIVENDA", "normativa": "Toda vivenda deberá contar coa dotación de instalacións mínimas obrigadas pola normativa de obrigado cumprimento.", "default": "SI"},
    {"id": "t5_r4", "table": 5, "row": 4, "section": "A.4.", "concepto": "A.4. CONDICIÓNS DOTACIONAIS DAS VIVENDAS", "title": "*REHABILITACIÓN: Nas actuacións de remodelación de vivendas será exixible a instalación de calefacción, a instalación dun sistema de ventilación e as infraestruturas de fogar dixital cando a vivenda existente xa conte con esas instalacións ou cando lles sexa obrigado executalas de acordo ao CTE ou a normativa sectorial.", "normativa": "SI", "default": "-"},
    {"id": "t5_r5", "table": 5, "row": 5, "section": "A.4.", "concepto": "A.4. CONDICIÓNS DOTACIONAIS DAS VIVENDAS", "title": "A.4.2.1. Cociña", "normativa": "Reserva de espazo, preinstalacións e tomas exixidas para: vertedoiro, lavalouzas, frigorífico, forno, cociña e espazo de almacenamento de lixos.", "default": "SI"},
    {"id": "t5_r6", "table": 5, "row": 6, "section": "A.4.", "concepto": "A.4. CONDICIÓNS DOTACIONAIS DAS VIVENDAS", "title": "A.4.2.1. Cociña", "normativa": "Zonas expostas á auga revestidas de material impermeable.", "default": "SI"},
    {"id": "t5_r7", "table": 5, "row": 7, "section": "A.4.", "concepto": "A.4. CONDICIÓNS DOTACIONAIS DAS VIVENDAS", "title": "Vivendas adaptadas co mobiliario de cociña de accesibilidade adaptable instalado.", "normativa": "SI", "default": "Non aplica"},
    {"id": "t5_r8", "table": 5, "row": 8, "section": "A.4.", "concepto": "A.4. CONDICIÓNS DOTACIONAIS DAS VIVENDAS", "title": "A.4.2.2. Cuarto de baño", "normativa": "Composto de bañeira/ducha, lavabo, inodoro e preinstalación para bidé.", "default": "SI"},
    {"id": "t5_r9", "table": 5, "row": 9, "section": "A.4.", "concepto": "A.4. CONDICIÓNS DOTACIONAIS DAS VIVENDAS", "title": "A.4.2.2. Cuarto de baño", "normativa": "Zonas expostas á auga revestidas de material impermeable.", "default": "SI"},
    {"id": "t5_r10", "table": 5, "row": 10, "section": "A.4.", "concepto": "A.4. CONDICIÓNS DOTACIONAIS DAS VIVENDAS", "title": "A.4.2.3. Cuarto de aseo", "normativa": "Cando sexa exixible de acordo ao número estancias da vivenda (≥ 4), contará, como mínimo, con lavabo e inodoro.", "default": "SI"},
    {"id": "t5_r11", "table": 5, "row": 11, "section": "A.4.", "concepto": "A.4. CONDICIÓNS DOTACIONAIS DAS VIVENDAS", "title": "A.4.2.3. Cuarto de aseo", "normativa": "Zonas expostas á auga revestidas de material impermeable.", "default": "SI"},
    {"id": "t5_r12", "table": 5, "row": 12, "section": "A.4.", "concepto": "A.4. CONDICIÓNS DOTACIONAIS DAS VIVENDAS", "title": "Preinstalación e tomas exixidas para: lavadora, lavadoiro e secadora.", "normativa": "SI", "default": "Non aplica"},
    {"id": "t5_r13", "table": 5, "row": 13, "section": "A.4.", "concepto": "A.4. CONDICIÓNS DOTACIONAIS DAS VIVENDAS", "title": "Revestimento en todos o seus paramentos de material impermeable ata altura de:", "normativa": "1,80 m.", "default": "Non aplica"},
]

DEFAULT_NHV_DATA = {
    "master": {
        "tipo_actuacion": "rehabilitacion",  # "obra_nova" | "rehabilitacion"
        "tipo_vivenda": "unifamiliar",      # "unifamiliar" | "colectiva"
        "num_estancias": 6,                 # 1 a 6
        "cocina_integrada": True,           # True | False
        "baixo_cuberta": False              # True | False
    },
    "dimens": {
        "sup_e1": "61,35 m2",
        "cad_e1": "3,30 x 3,30 m",
        "sup_e2": "32,25 m2",
        "cad_e2": "2,60 x 2,60 m",
        "sup_e3": "12,90 m2",
        "cad_e3": "2,20 x 2,20 m",
        "sup_e4": "12,20 m2",
        "cad_e4": "2,20 x 2,20 m",
        "sup_e5": "10,77 m2",
        "cad_e5": "2,20 x 2,20 m",
        "sup_e6": "9,86 m2",
        "cad_e6": "2,20 x 2,20 m",
        "sup_cocina": "10,00 m2",
        "mesado_cocina": "3,77 m",
        "sup_bano": "5,49 m2",
        "sup_aseo": "3,55 m2",
        "sup_almacen": "6,03 m2",
        "fondo_almacen": "0,60 m",
        "ancho_portas": "0,80 m",
        "alto_portas": "2,03 m"
    },
    "items": {
        "t0_r1": "SI",
        "t0_r2": "SI",
        "t0_r3": "Non aplica",
        "t0_r4": "Non aplica",
        "t0_r5": "Non aplica",
        "t0_r6": "Non aplica",
        "t0_r7": "SI",
        "t1_r0": "Existente (Rehab.)",
        "t1_r1": "Existente (Rehab.)",
        "t1_r2": "Existente (Rehab.)",
        "t1_r3": "NON",
        "t1_r4": "Non aplica",
        "t1_r5": "≥ 2 m",
        "t1_r6": "Non aplica",
        "t1_r7": "Non aplica",
        "t1_r8": "Non aplica",
        "t1_r9": "Non aplica",
        "t1_r10": "Non aplica",
        "t1_r11": "Existente (Rehab.)",
        "t1_r12": "Existente (Rehab.)",
        "t1_r13": "Existente (Rehab.)",
        "t1_r14": "Existente (Rehab.)",
        "t1_r15": "Existente (Rehab.)",
        "t1_r16": "Existente (Rehab.)",
        "t1_r17": "SI",
        "t1_r18": "NON",
        "t1_r19": "SI",
        "t1_r20": "Non aplica",
        "t1_r21": "Non aplica",
        "t1_r22": "SI",
        "t1_r23": "Non aplica",
        "t2_r0": "NON",
        "t2_r1": "SI",
        "t2_r2": "SI",
        "t2_r3": "SI",
        "t2_r4": "SI",
        "t2_r5": "Baño dende zona común",
        "t2_r6": "Non aplica",
        "t2_r7": "Non aplica",
        "t2_r8": "SI",
        "t2_r9": "≥ 2,50 m",
        "t2_r10": "≤ 30 %",
        "t2_r11": "≥ 2,20 m",
        "t2_r12": "h existente cumpre",
        "t2_r13": "h existente cumpre",
        "t2_r14": "",
        "t2_r15": "Non aplica",
        "t2_r16": "Non aplica",
        "t2_r17": "Non aplica",
        "t2_r19": "Estancia maior, E1",
        "t2_r20": "Non aplica",
        "t2_r21": "Non aplica",
        "t2_r22": "Non aplica",
        "t2_r23": "Non aplica",
        "t2_r24": "Non aplica",
        "t2_r25": "61,35 m2",
        "t2_r26": "Non aplica",
        "t2_r27": "3,30 x 3,30 m",
        "t3_r0": "0,00 m2",
        "t3_r1": "4,56 m",
        "t3_r2": "Estancia E2",
        "t3_r3": "32,25 m2",
        "t3_r4": "2,60x2,60 m",
        "t3_r5": "0,00 m2",
        "t3_r6": "5,27 m",
        "t3_r7": "0%",
        "t3_r8": "Estancia E3",
        "t3_r9": "12,90 m2",
        "t3_r10": "2,20x2,20 m",
        "t3_r11": "0,00 m2",
        "t3_r12": "2,79 m",
        "t3_r13": "0%",
        "t3_r14": "Estancia E4",
        "t3_r15": "12,20 m2",
        "t3_r16": "2,20x2,20 m",
        "t3_r17": "0,00 m2",
        "t3_r18": "2,81 m",
        "t3_r19": "0%",
        "t3_r20": "Estancia E5",
        "t3_r21": "Non aplica",
        "t3_r22": "10,77 m2",
        "t3_r23": "2,20x2,20 m",
        "t3_r24": "0,00 m2",
        "t3_r25": "2,79 m",
        "t3_r26": "0%",
        "t3_r27": "Estancia E6",
        "t3_r28": "9,86 m2",
        "t3_r29": "2,20x2,20 m",
        "t3_r30": "0,00 m2",
        "t3_r31": "2,79 m",
        "t3_r32": "0%",
        "t4_r0": "Outras estancias",
        "t4_r1": "SI",
        "t4_r2": "Non aplica",
        "t4_r4": "Cociña",
        "t4_r5": "Non aplica",
        "t4_r6": "Non aplica",
        "t4_r7": "Non aplica",
        "t4_r8": "Non aplica",
        "t4_r9": "Non aplica",
        "t4_r10": "",
        "t4_r11": "61,35 m2 (>25+10+2)",
        "t4_r12": "7,89 m2",
        "t4_r13": "1,91 m",
        "t4_r14": "Non aplica",
        "t4_r15": "3,77 m",
        "t4_r16": "0,90 m",
        "t4_r17": "2,20x2,20 m",
        "t4_r18": "0,00 m2",
        "t4_r19": "0%",
        "t4_r20": "Cuarto de baño",
        "t4_r21": "5,49 m2",
        "t4_r22": "1,97 m",
        "t4_r23": "0,00 m2",
        "t4_r24": "SI",
        "t4_r25": "Non aplica",
        "t4_r26": "Cuarto de aseo",
        "t4_r27": "3,55 m2",
        "t4_r28": "1,80 m",
        "t4_r29": "0,00 m2",
        "t4_r30": "Lavadoiro",
        "t4_r31": "m2",
        "t4_r32": "m",
        "t4_r33": "Non aplica",
        "t4_r34": "Non aplica",
        "t4_r35": "Non aplica",
        "t4_r36": "Non aplica",
        "t4_r37": "Non aplica",
        "t4_r38": "Tendal",
        "t4_r39": "Non aplica",
        "t4_r40": "Non aplica",
        "t4_r41": "Non aplica",
        "t4_r42": "Non aplica",
        "t4_r43": "Non aplica",
        "t4_r44": "Non aplica",
        "t4_r45": "Non aplica",
        "t4_r46": "Non aplica",
        "t4_r47": "Non aplica",
        "t4_r48": "Non aplica",
        "t4_r49": "Non aplica",
        "t4_r50": "Non aplica",
        "t4_r51": "Non aplica",
        "t4_r52": "SI",
        "t4_r53": "Espazos de almacenamento",
        "t4_r54": "Non aplica",
        "t4_r55": "Non aplica",
        "t4_r56": "Non aplica",
        "t4_r57": "Non aplica",
        "t4_r58": "Non aplica",
        "t4_r59": "6,03 m2",
        "t4_r60": "2,20 m",
        "t4_r61": "0,60 m",
        "t4_r62": "0,67 m2",
        "t4_r63": "Dormitorios",
        "t4_r64": "Dormitorios",
        "t4_r66": "Espazo de acceso interior da vivenda",
        "t4_r67": "1,50x1,50 m",
        "t4_r68": "Corredores e zonas de acceso interiores as pezas",
        "t4_r69": "1,60 m",
        "t4_r70": "1,60 m",
        "t5_r1": "0,80 m",
        "t5_r2": "2,03 m",
        "t5_r3": "SI",
        "t5_r4": "-",
        "t5_r5": "SI",
        "t5_r6": "SI",
        "t5_r7": "Non aplica",
        "t5_r8": "SI",
        "t5_r9": "SI",
        "t5_r10": "SI",
        "t5_r11": "SI",
        "t5_r12": "Non aplica",
        "t5_r13": "Non aplica"
}
}


def compute_nhv_defaults(master_params, custom_dimens=None):
    """
    Aplica as regras do Decreto 29/2010 segundo os 5 parámetros mestres
    e xera os valores recomendados para as 173 celas de 'PROXECTO'.
    """
    res = dict(DEFAULT_NHV_DATA["items"])
    act = master_params.get("tipo_actuacion", "rehabilitacion")
    viv = master_params.get("tipo_vivenda", "unifamiliar")
    nest = int(master_params.get("num_estancias", 6))
    cint = bool(master_params.get("cocina_integrada", True))
    bcub = bool(master_params.get("baixo_cuberta", False))

    dimens = dict(DEFAULT_NHV_DATA["dimens"])
    if custom_dimens:
        dimens.update(custom_dimens)

    # 1. Regras de Rehabilitación vs Obra Nova en Táboa 1 (Iluminación e ventilación)
    if act == "rehabilitacion":
        for r_id in ["t1_r0", "t1_r1", "t1_r2", "t1_r11", "t1_r12", "t1_r13", "t1_r14", "t1_r15", "t1_r16"]:
            res[r_id] = "Existente (Rehab.)"
        res["t1_r17"] = "SI"
        res["t1_r18"] = "NON"
        res["t1_r19"] = "SI"
        res["t2_r12"] = "h existente cumpre"
        res["t2_r13"] = "h existente cumpre"
        res["t4_r52"] = "SI"
        res["t5_r4"] = "-"
    else:
        res["t1_r0"] = "Cumpre (1/8)"
        res["t1_r1"] = "1,10 m"
        res["t1_r2"] = "0,50 m"
        res["t1_r11"] = "Cumpre"
        res["t1_r12"] = "Non aplica"
        res["t1_r13"] = "Non aplica"
        res["t1_r14"] = "Non aplica"
        res["t1_r15"] = "Non aplica"
        res["t1_r16"] = "Cumpre (1/3)"
        res["t1_r17"] = "Non aplica"
        res["t1_r18"] = "Non aplica"
        res["t1_r19"] = "Non aplica"
        res["t2_r12"] = "Non aplica"
        res["t2_r13"] = "Non aplica"
        res["t4_r52"] = "Non aplica"
        res["t5_r4"] = "SI"

    # 2. Vivenda Unifamiliar vs Colectiva
    if viv == "unifamiliar":
        res["t0_r6"] = "Non aplica"
        res["t4_r35"] = "Non aplica"
        res["t4_r51"] = "Non aplica"
    else:
        res["t0_r6"] = "Non aplica"
        res["t4_r35"] = "Non aplica"
        res["t4_r51"] = "Non aplica"

    # 3. Baixo cuberta
    if not bcub:
        for r_id in ["t2_r14", "t2_r15", "t2_r16", "t2_r17"]:
            res[r_id] = "Non aplica"
    else:
        res["t2_r14"] = "SI"
        res["t2_r15"] = "Cumpre"
        res["t2_r16"] = ">= 2,20 m"
        res["t2_r17"] = ">= 1,80 m"

    # 4. Número de estancias (E1 en Táboa 2)
    # t2_r20 (=1), t2_r21 (=2), t2_r22 (=3), t2_r23 (=4), t2_r24 (=5), t2_r25 (>5)
    e1_map = {1: "t2_r20", 2: "t2_r21", 3: "t2_r22", 4: "t2_r23", 5: "t2_r24", 6: "t2_r25"}
    for n, r_id in e1_map.items():
        if (nest == n) or (nest >= 6 and n == 6):
            res[r_id] = dimens.get("sup_e1", "25,00 m2")
        else:
            res[r_id] = "Non aplica"

    # Estancias secundarias en Táboa 3
    # E2 aplica sempre se nest >= 2
    if nest >= 2:
        res["t3_r3"] = dimens.get("sup_e2", "32,25 m2")
        res["t3_r4"] = dimens.get("cad_e2", "2,60x2,60 m")
    else:
        res["t3_r3"] = "Non aplica"
        res["t3_r4"] = "Non aplica"

    # E3 aplica se nest >= 3
    if nest >= 3:
        res["t3_r9"] = dimens.get("sup_e3", "12,90 m2")
        res["t3_r10"] = dimens.get("cad_e3", "2,20x2,20 m")
    else:
        res["t3_r9"] = "Non aplica"
        res["t3_r10"] = "Non aplica"

    # E4 aplica se nest >= 4
    if nest >= 4:
        res["t3_r15"] = dimens.get("sup_e4", "12,20 m2")
        res["t3_r16"] = dimens.get("cad_e4", "2,20x2,20 m")
    else:
        res["t3_r15"] = "Non aplica"
        res["t3_r16"] = "Non aplica"

    # E5 aplica se nest >= 5
    if nest == 5:
        res["t3_r21"] = dimens.get("sup_e5", "6,00 m2")
        res["t3_r22"] = "Non aplica"
    elif nest > 5:
        res["t3_r21"] = "Non aplica"
        res["t3_r22"] = dimens.get("sup_e5", "10,77 m2")
        res["t3_r23"] = dimens.get("cad_e5", "2,20x2,20 m")
    else:
        res["t3_r21"] = "Non aplica"
        res["t3_r22"] = "Non aplica"
        res["t3_r23"] = "Non aplica"

    # E6 aplica se nest >= 6
    if nest >= 6:
        res["t3_r28"] = dimens.get("sup_e6", "9,86 m2")
        res["t3_r29"] = dimens.get("cad_e6", "2,20x2,20 m")
    else:
        res["t3_r28"] = "Non aplica"
        res["t3_r29"] = "Non aplica"

    # Cociña en Táboa 4
    coc_map = {1: "t4_r5", 2: "t4_r6", 3: "t4_r7", 4: "t4_r8", 5: "t4_r9", 6: "t4_r10"}
    for n, r_id in coc_map.items():
        if (nest == n) or (nest >= 6 and n == 6):
            res[r_id] = dimens.get("sup_cocina", "10,00 m2")
        else:
            res[r_id] = "Non aplica"

    if cint:
        res["t4_r11"] = f"{dimens.get('sup_e1', '61,35 m2')} (Cumpre)"
        res["t4_r12"] = "7,89 m2"
    else:
        res["t4_r11"] = "Non aplica"
        res["t4_r12"] = "Non aplica"

    res["t4_r15"] = dimens.get("mesado_cocina", "3,77 m")

    # Baño e aseo en Táboa 4
    res["t4_r21"] = dimens.get("sup_bano", "5,49 m2")
    if nest >= 4:
        res["t4_r27"] = dimens.get("sup_aseo", "3,55 m2")
        res["t5_r10"] = "SI"
    else:
        res["t4_r27"] = "Non aplica"
        res["t5_r10"] = "Non aplica"

    # Almacenamento en Táboa 4
    alm_map = {1: "t4_r54", 2: "t4_r55", 3: "t4_r56", 4: "t4_r57", 5: "t4_r58", 6: "t4_r59"}
    for n, r_id in alm_map.items():
        if (nest == n) or (nest >= 6 and n == 6):
            res[r_id] = dimens.get("sup_almacen", "6,03 m2")
        else:
            res[r_id] = "Non aplica"

    res["t4_r61"] = dimens.get("fondo_almacen", "0,60 m")

    # Portas en Táboa 5
    res["t5_r1"] = dimens.get("ancho_portas", "0,80 m")
    res["t5_r2"] = dimens.get("alto_portas", "2,03 m")

    return res


def _set_cell_text_preserving_style(cell, text):
    """Substitúe o texto da celda conservando estilo, fonte, tamaño e aliñación."""
    if not cell.paragraphs:
        cell.text = str(text).strip() if text is not None else ""
        return
    p = cell.paragraphs[0]
    clean_text = str(text).strip() if text is not None else ""
    if p.runs:
        p.runs[0].text = clean_text
        for r in p.runs[1:]:
            r.text = ""
    else:
        p.text = clean_text
    # Limpar parágrafos adicionais se os houbese na mesma cela
    for p_extra in cell.paragraphs[1:]:
        p_extra.text = ""


def generate_nhv(project_data, output_stream=None, output_path=None):
    """
    Xera o documento .docx de xustificación de Habitabilidade NHV
    a partir dos datos do proxecto e da plantilla NHV.docx.
    """
    template_path, _ = get_default_nhv_paths()
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Non se atopou o arquivo de plantilla NHV.docx en: {template_path}")

    doc = docx.Document(template_path)

    # 1. Obter datos mestres e items
    nhv_dict = project_data.get("nhv_data", {})
    if not nhv_dict:
        # Intentar se foi pasado directamente
        nhv_dict = project_data

    master = nhv_dict.get("master", DEFAULT_NHV_DATA["master"])
    dimens = nhv_dict.get("dimens", DEFAULT_NHV_DATA["dimens"])
    items = nhv_dict.get("items")
    if not items:
        items = compute_nhv_defaults(master, dimens)

    # 2. Actualizar tódalas celas da columna PROXECTO nas táboas 0 a 5
    for cat in NHV_CATALOG:
        t_idx = cat["table"]
        r_idx = cat["row"]
        row_id = cat["id"]
        
        if row_id in items:
            val = items[row_id]
            if t_idx < len(doc.tables):
                table = doc.tables[t_idx]
                if r_idx < len(table.rows):
                    row = table.rows[r_idx]
                    target_cell = row.cells[-1]
                    _set_cell_text_preserving_style(target_cell, val)

    # 3. Actualizar data e poboación no parágrafo P14 ('Ferrol, xullo 2026')
    poboacion = project_data.get("poboacion") or project_data.get("concello") or "Ferrol"
    data_prox = project_data.get("data_proxecto") or project_data.get("data") or "xullo 2026"
    fecha_texto = f"{poboacion}, {data_prox}"

    for p in doc.paragraphs:
        txt = p.text.strip()
        if "Ferrol, xullo 2026" in txt or ("," in txt and ("202" in txt or "201" in txt)):
            if p.runs:
                p.runs[0].text = fecha_texto
                for r in p.runs[1:]:
                    r.text = ""
            else:
                p.text = fecha_texto
            break

    # 4. Actualizar cadro de firmas en Táboa 6
    if len(doc.tables) > 6:
        t6 = doc.tables[6]
        autor = project_data.get("autor_proxecto") or project_data.get("arquitecto") or "Estudio Anta Arquitectos S.L.P."
        colexiado = project_data.get("num_colexiado") or "20.039"
        colexio = project_data.get("colexio_profesional") or "COAG"
        
        if len(t6.rows) > 0 and len(t6.rows[0].cells) > 0:
            _set_cell_text_preserving_style(t6.rows[0].cells[0], f"Fdo. {autor}")
        if len(t6.rows) > 1 and len(t6.rows[1].cells) > 0:
            _set_cell_text_preserving_style(t6.rows[1].cells[0], autor)
            if len(t6.rows[1].cells) >= 3:
                _set_cell_text_preserving_style(t6.rows[1].cells[2], f"Nº {colexio} - {colexiado}")

    # 5. Gardar no destino
    if output_stream is not None:
        doc.save(output_stream)
        output_stream.seek(0)
        return output_stream
    elif output_path is not None:
        doc.save(output_path)
        return output_path
    else:
        stream = io.BytesIO()
        doc.save(stream)
        stream.seek(0)
        return stream
