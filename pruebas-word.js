// ══════════════════════════════════════════════════════════════
// PRUEBAS EN WORD (.docx) · CEIP Capitulaciones de Santa Fe
// Lectura, Cálculo y Resolución de problemas.
// Cada prueba se descarga en dos documentos separados:
//   · Cuadernillo del alumnado (solo lo que se entrega al niño/a)
//   · Cuadernillo del maestro/a (guía, baremos, soluciones y registro)
// Cada hoja es una sección con su cabecera y el logo del colegio.
// Se abre con Word, LibreOffice y Google Docs.
// ══════════════════════════════════════════════════════════════

// Banco de cálculo: Parte 1 = cálculo mental (10 ítems) · Parte 2 = operaciones escritas (6)
const PRUEBAS_CALC = {"1º": {"1T": {"contenidos": "Sumas y restas hasta 10; anterior y posterior.", "tiempo": 5, "mental": [{"q": "3 + 4 =", "r": "7"}, {"q": "3 + 7 =", "r": "10"}, {"q": "1 + 6 =", "r": "7"}, {"q": "5 + 4 =", "r": "9"}, {"q": "4 + 6 =", "r": "10"}, {"q": "4 − 3 =", "r": "1"}, {"q": "7 − 4 =", "r": "3"}, {"q": "6 − 5 =", "r": "1"}, {"q": "El número que va después del 10 es", "r": "11"}, {"q": "El número que va antes del 5 es", "r": "4"}], "escrito": [{"t": "v", "a": "6", "s": "+", "b": "2", "r": "8"}, {"t": "v", "a": "4", "s": "+", "b": "1", "r": "5"}, {"t": "v", "a": "3", "s": "+", "b": "3", "r": "6"}, {"t": "v", "a": "4", "s": "−", "b": "1", "r": "3"}, {"t": "v", "a": "6", "s": "−", "b": "3", "r": "3"}, {"t": "v", "a": "4", "s": "−", "b": "3", "r": "1"}]}, "2T": {"contenidos": "Números hasta 50; sumas y restas sin llevar.", "tiempo": 5, "mental": [{"q": "10 + 6 =", "r": "16"}, {"q": "40 + 4 =", "r": "44"}, {"q": "5 + 3 =", "r": "8"}, {"q": "2 + 3 =", "r": "5"}, {"q": "18 − 1 =", "r": "17"}, {"q": "28 − 3 =", "r": "25"}, {"q": "10 + 10 =", "r": "20"}, {"q": "38 + 10 =", "r": "48"}, {"q": "3 − 2 =", "r": "1"}, {"q": "7 − 3 =", "r": "4"}], "escrito": [{"t": "v", "a": "14", "s": "+", "b": "1", "r": "15"}, {"t": "v", "a": "32", "s": "+", "b": "3", "r": "35"}, {"t": "v", "a": "34", "s": "+", "b": "1", "r": "35"}, {"t": "v", "a": "25", "s": "−", "b": "3", "r": "22"}, {"t": "v", "a": "27", "s": "−", "b": "5", "r": "22"}, {"t": "v", "a": "45", "s": "−", "b": "2", "r": "43"}]}, "3T": {"contenidos": "Números hasta 99; sumas y restas sin llevar; decenas; dobles.", "tiempo": 5, "mental": [{"q": "77 + 2 =", "r": "79"}, {"q": "80 + 8 =", "r": "88"}, {"q": "37 − 6 =", "r": "31"}, {"q": "33 − 1 =", "r": "32"}, {"q": "77 + 10 =", "r": "87"}, {"q": "49 − 10 =", "r": "39"}, {"q": "20 + 30 =", "r": "50"}, {"q": "30 − 10 =", "r": "20"}, {"q": "El doble de 4 es", "r": "8"}, {"q": "El doble de 2 es", "r": "4"}], "escrito": [{"t": "v", "a": "47", "s": "+", "b": "30", "r": "77"}, {"t": "v", "a": "14", "s": "+", "b": "20", "r": "34"}, {"t": "v", "a": "51", "s": "+", "b": "18", "r": "69"}, {"t": "v", "a": "91", "s": "−", "b": "31", "r": "60"}, {"t": "v", "a": "37", "s": "−", "b": "32", "r": "5"}, {"t": "v", "a": "77", "s": "−", "b": "21", "r": "56"}]}}, "2º": {"1T": {"contenidos": "Sumas y restas sin llevar; decenas; dobles; sumas con llevada de 2 + 1 cifras.", "tiempo": 5, "mental": [{"q": "34 + 4 =", "r": "38"}, {"q": "68 − 6 =", "r": "62"}, {"q": "10 + 10 =", "r": "20"}, {"q": "59 − 10 =", "r": "49"}, {"q": "70 + 20 =", "r": "90"}, {"q": "40 − 30 =", "r": "10"}, {"q": "El doble de 9 es", "r": "18"}, {"q": "El doble de 10 es", "r": "20"}, {"q": "La mitad de 10 es", "r": "5"}, {"q": "14 + 4 =", "r": "18"}], "escrito": [{"t": "v", "a": "40", "s": "+", "b": "19", "r": "59"}, {"t": "v", "a": "62", "s": "+", "b": "20", "r": "82"}, {"t": "v", "a": "99", "s": "−", "b": "28", "r": "71"}, {"t": "v", "a": "99", "s": "−", "b": "20", "r": "79"}, {"t": "v", "a": "47", "s": "+", "b": "5", "r": "52"}, {"t": "v", "a": "49", "s": "+", "b": "7", "r": "56"}]}, "2T": {"contenidos": "Sumas con llevada; restas sin llevar de 3 cifras; tablas del 2, 5 y 10.", "tiempo": 5, "mental": [{"q": "2 × 9 =", "r": "18"}, {"q": "10 × 10 =", "r": "100"}, {"q": "5 × 3 =", "r": "15"}, {"q": "2 × 7 =", "r": "14"}, {"q": "11 + 27 =", "r": "38"}, {"q": "58 + 11 =", "r": "69"}, {"q": "20 + 10 =", "r": "30"}, {"q": "70 − 10 =", "r": "60"}, {"q": "La mitad de 18 es", "r": "9"}, {"q": "El doble de 5 es", "r": "10"}], "escrito": [{"t": "v", "a": "46", "s": "+", "b": "27", "r": "73"}, {"t": "v", "a": "52", "s": "+", "b": "18", "r": "70"}, {"t": "v", "a": "26", "s": "+", "b": "18", "r": "44"}, {"t": "v", "a": "280", "s": "−", "b": "50", "r": "230"}, {"t": "v", "a": "614", "s": "−", "b": "13", "r": "601"}, {"t": "v", "a": "277", "s": "−", "b": "74", "r": "203"}]}, "3T": {"contenidos": "Restas con llevada; tablas del 2 al 5 y del 10; multiplicación sin llevar.", "tiempo": 5, "mental": [{"q": "5 × 7 =", "r": "35"}, {"q": "2 × 6 =", "r": "12"}, {"q": "2 × 5 =", "r": "10"}, {"q": "3 × 5 =", "r": "15"}, {"q": "4 × 9 =", "r": "36"}, {"q": "3 × 10 =", "r": "30"}, {"q": "55 + 34 =", "r": "89"}, {"q": "50 − 30 =", "r": "20"}, {"q": "300 + 100 =", "r": "400"}, {"q": "La mitad de 28 es", "r": "14"}], "escrito": [{"t": "v", "a": "96", "s": "−", "b": "39", "r": "57"}, {"t": "v", "a": "51", "s": "−", "b": "38", "r": "13"}, {"t": "v", "a": "84", "s": "−", "b": "38", "r": "46"}, {"t": "v", "a": "33", "s": "×", "b": "2", "r": "66"}, {"t": "v", "a": "32", "s": "×", "b": "2", "r": "64"}, {"t": "v", "a": "20", "s": "×", "b": "3", "r": "60"}]}}, "3º": {"1T": {"contenidos": "Sumas y restas con llevada de 3 cifras; multiplicación por 1 cifra; tablas del 2 al 6.", "tiempo": 4, "mental": [{"q": "2 × 5 =", "r": "10"}, {"q": "6 × 8 =", "r": "48"}, {"q": "4 × 8 =", "r": "32"}, {"q": "6 × 5 =", "r": "30"}, {"q": "6 × 7 =", "r": "42"}, {"q": "300 + 400 =", "r": "700"}, {"q": "600 − 200 =", "r": "400"}, {"q": "214 + 100 =", "r": "314"}, {"q": "50 + 12 =", "r": "62"}, {"q": "45 × 10 =", "r": "450"}], "escrito": [{"t": "v", "a": "580", "s": "+", "b": "280", "r": "860"}, {"t": "v", "a": "359", "s": "+", "b": "213", "r": "572"}, {"t": "v", "a": "932", "s": "−", "b": "346", "r": "586"}, {"t": "v", "a": "973", "s": "−", "b": "138", "r": "835"}, {"t": "v", "a": "18", "s": "×", "b": "4", "r": "72"}, {"t": "v", "a": "42", "s": "×", "b": "4", "r": "168"}]}, "2T": {"contenidos": "Multiplicación de 3 cifras por 1 cifra; restas con llevada; tablas hasta el 9; × 10 y × 100.", "tiempo": 4, "mental": [{"q": "6 × 3 =", "r": "18"}, {"q": "9 × 7 =", "r": "63"}, {"q": "9 × 4 =", "r": "36"}, {"q": "8 × 8 =", "r": "64"}, {"q": "8 × 10 =", "r": "80"}, {"q": "4 × 7 =", "r": "28"}, {"q": "57 × 10 =", "r": "570"}, {"q": "4 × 100 =", "r": "400"}, {"q": "El doble de 37 es", "r": "74"}, {"q": "La mitad de 80 es", "r": "40"}], "escrito": [{"t": "v", "a": "178", "s": "×", "b": "3", "r": "534"}, {"t": "v", "a": "166", "s": "×", "b": "4", "r": "664"}, {"t": "v", "a": "173", "s": "×", "b": "2", "r": "346"}, {"t": "v", "a": "668", "s": "−", "b": "376", "r": "292"}, {"t": "v", "a": "914", "s": "−", "b": "345", "r": "569"}, {"t": "v", "a": "189", "s": "+", "b": "137", "r": "326"}]}, "3T": {"contenidos": "División entre 1 cifra; multiplicación por 1 cifra; tablas e inversas.", "tiempo": 4, "mental": [{"q": "7 × 4 =", "r": "28"}, {"q": "7 × 8 =", "r": "56"}, {"q": "6 × 6 =", "r": "36"}, {"q": "8 × 4 =", "r": "32"}, {"q": "16 : 8 =", "r": "2"}, {"q": "56 : 8 =", "r": "7"}, {"q": "56 : 7 =", "r": "8"}, {"q": "6 : 2 =", "r": "3"}, {"q": "13 × 100 =", "r": "1300"}, {"q": "50 : 10 =", "r": "5"}], "escrito": [{"t": "d", "a": "84", "b": "4", "r": "c = 21, r = 0"}, {"t": "d", "a": "87", "b": "5", "r": "c = 17, r = 2"}, {"t": "d", "a": "342", "b": "3", "r": "c = 114, r = 0"}, {"t": "v", "a": "416", "s": "×", "b": "5", "r": "2080"}, {"t": "v", "a": "454", "s": "×", "b": "4", "r": "1816"}, {"t": "v", "a": "481", "s": "×", "b": "7", "r": "3367"}]}}, "4º": {"1T": {"contenidos": "Multiplicación por 2 cifras; división entre 1 cifra; tablas, × y : por 10, 100 y 1000.", "tiempo": 4, "mental": [{"q": "7 × 7 =", "r": "49"}, {"q": "9 × 9 =", "r": "81"}, {"q": "9 × 6 =", "r": "54"}, {"q": "7 × 4 =", "r": "28"}, {"q": "63 : 7 =", "r": "9"}, {"q": "49 : 7 =", "r": "7"}, {"q": "81 : 9 =", "r": "9"}, {"q": "33 × 100 =", "r": "3300"}, {"q": "2 × 1000 =", "r": "2000"}, {"q": "630 : 10 =", "r": "63"}], "escrito": [{"t": "v", "a": "248", "s": "×", "b": "26", "r": "6448"}, {"t": "v", "a": "145", "s": "×", "b": "29", "r": "4205"}, {"t": "v", "a": "214", "s": "×", "b": "20", "r": "4280"}, {"t": "d", "a": "857", "b": "9", "r": "c = 95, r = 2"}, {"t": "d", "a": "409", "b": "6", "r": "c = 68, r = 1"}, {"t": "d", "a": "821", "b": "6", "r": "c = 136, r = 5"}]}, "2T": {"contenidos": "División entre 1 y 2 cifras; multiplicación por 2 cifras; cálculo mental de 2 cifras por 1.", "tiempo": 4, "mental": [{"q": "19 × 4 =", "r": "76"}, {"q": "13 × 2 =", "r": "26"}, {"q": "21 × 4 =", "r": "84"}, {"q": "63 : 7 =", "r": "9"}, {"q": "21 : 7 =", "r": "3"}, {"q": "54 : 9 =", "r": "6"}, {"q": "700 : 100 =", "r": "7"}, {"q": "El doble de 21 es", "r": "42"}, {"q": "La mitad de 60 es", "r": "30"}, {"q": "50 × 8 =", "r": "400"}], "escrito": [{"t": "v", "a": "498", "s": "×", "b": "27", "r": "13446"}, {"t": "v", "a": "307", "s": "×", "b": "39", "r": "11973"}, {"t": "d", "a": "4200", "b": "5", "r": "c = 840, r = 0"}, {"t": "d", "a": "1512", "b": "9", "r": "c = 168, r = 0"}, {"t": "d", "a": "90", "b": "18", "r": "c = 5, r = 0"}, {"t": "d", "a": "225", "b": "37", "r": "c = 6, r = 3"}]}, "3T": {"contenidos": "División entre 2 cifras; sumas y restas con decimales; multiplicación de decimal por entero.", "tiempo": 4, "mental": [{"q": "5,00 € + 1,50 € =", "r": "6,50 €"}, {"q": "5,00 € + 3,50 € =", "r": "8,50 €"}, {"q": "4,00 € + 2,50 € =", "r": "6,50 €"}, {"q": "48 : 6 =", "r": "8"}, {"q": "40 : 8 =", "r": "5"}, {"q": "15 × 4 =", "r": "60"}, {"q": "50 × 8 =", "r": "400"}, {"q": "1/4 de 24 =", "r": "6"}, {"q": "800 : 100 =", "r": "8"}, {"q": "18 × 1000 =", "r": "18000"}], "escrito": [{"t": "d", "a": "569", "b": "28", "r": "c = 20, r = 9"}, {"t": "d", "a": "588", "b": "22", "r": "c = 26, r = 16"}, {"t": "v", "a": "20,91", "s": "+", "b": "4,96", "r": "25,87"}, {"t": "v", "a": "41,19", "s": "+", "b": "7,04", "r": "48,23"}, {"t": "v", "a": "89,89", "s": "−", "b": "3,30", "r": "86,59"}, {"t": "v", "a": "5,4", "s": "×", "b": "8", "r": "43,2"}]}}, "5º": {"1T": {"contenidos": "División entre 2 cifras; multiplicación por 3 cifras; restas con llevadas; operaciones combinadas sencillas.", "tiempo": 4, "mental": [{"q": "16 × 1000 =", "r": "16000"}, {"q": "2100 : 100 =", "r": "21"}, {"q": "58000 : 1000 =", "r": "58"}, {"q": "24 : 8 =", "r": "3"}, {"q": "32 : 8 =", "r": "4"}, {"q": "13 × 2 =", "r": "26"}, {"q": "25 × 3 =", "r": "75"}, {"q": "8 + 5 × 5 =", "r": "33"}, {"q": "4 + 4 × 2 =", "r": "12"}, {"q": "50 × 2 =", "r": "100"}], "escrito": [{"t": "d", "a": "8704", "b": "30", "r": "c = 290, r = 4"}, {"t": "d", "a": "4028", "b": "39", "r": "c = 103, r = 11"}, {"t": "d", "a": "6687", "b": "60", "r": "c = 111, r = 27"}, {"t": "v", "a": "757", "s": "×", "b": "289", "r": "218773"}, {"t": "v", "a": "4860", "s": "×", "b": "47", "r": "228420"}, {"t": "v", "a": "61524", "s": "−", "b": "18789", "r": "42735"}]}, "2T": {"contenidos": "Números decimales: sumar, restar, multiplicar y dividir entre un entero; fracción de un número.", "tiempo": 4, "mental": [{"q": "3,61 × 10 =", "r": "36,1"}, {"q": "7,88 × 100 =", "r": "788"}, {"q": "14,3 : 10 =", "r": "1,43"}, {"q": "1,6 + 3,9 =", "r": "5,5"}, {"q": "2,6 + 2,2 =", "r": "4,8"}, {"q": "1/3 de 21 =", "r": "7"}, {"q": "1/10 de 50 =", "r": "5"}, {"q": "1/5 de 10 =", "r": "2"}, {"q": "El 50 % de 90 es", "r": "45"}, {"q": "La mitad de 216 es", "r": "108"}], "escrito": [{"t": "v", "a": "64,38", "s": "+", "b": "6,84", "r": "71,22"}, {"t": "v", "a": "59,93", "s": "−", "b": "1,96", "r": "57,97"}, {"t": "v", "a": "9,38", "s": "×", "b": "8", "r": "75,04"}, {"t": "v", "a": "2,7", "s": "×", "b": "4,6", "r": "12,42"}, {"t": "d", "a": "8,92", "b": "2", "r": "c = 4,46"}, {"t": "d", "a": "29,8", "b": "5", "r": "c = 5,96"}]}, "3T": {"contenidos": "Decimales y porcentajes sencillos; fracción de un número; operaciones combinadas.", "tiempo": 4, "mental": [{"q": "El 10 % de 400 es", "r": "40"}, {"q": "El 10 % de 500 es", "r": "50"}, {"q": "El 50 % de 100 es", "r": "50"}, {"q": "3/4 de 24 =", "r": "18"}, {"q": "4/5 de 50 =", "r": "40"}, {"q": "9,69 × 100 =", "r": "969"}, {"q": "50,5 : 100 =", "r": "0,505"}, {"q": "2 + 5 × 3 =", "r": "17"}, {"q": "(4 + 3) × 2 =", "r": "14"}, {"q": "1/10 + 4/10 =", "r": "5/10"}], "escrito": [{"t": "d", "a": "42,35", "b": "7", "r": "c = 6,05"}, {"t": "d", "a": "27,85", "b": "5", "r": "c = 5,57"}, {"t": "d", "a": "6427", "b": "58", "r": "c = 110, r = 47"}, {"t": "v", "a": "4,8", "s": "×", "b": "2", "r": "9,6"}, {"t": "v", "a": "38,62", "s": "−", "b": "7,38", "r": "31,24"}, {"t": "v", "a": "82,21", "s": "+", "b": "5,28", "r": "87,49"}]}}, "6º": {"1T": {"contenidos": "Operaciones combinadas con paréntesis; potencias; decimales por la unidad seguida de ceros; división entre 2 cifras.", "tiempo": 4, "mental": [{"q": "3³ =", "r": "27"}, {"q": "4² =", "r": "16"}, {"q": "5² =", "r": "25"}, {"q": "(3 + 3) × 2 =", "r": "12"}, {"q": "(6 + 4) × 3 =", "r": "30"}, {"q": "9 + 4 × 5 =", "r": "29"}, {"q": "4,99 × 100 =", "r": "499"}, {"q": "68,9 : 10 =", "r": "6,89"}, {"q": "83 : 100 =", "r": "0,83"}, {"q": "El 50 % de 90 es", "r": "45"}], "escrito": [{"t": "h", "q": "135 − 4 × 7 =", "r": "107"}, {"t": "h", "q": "115 − 6 × 9 =", "r": "61"}, {"t": "h", "q": "8 × (25 − 10) + 6 =", "r": "126"}, {"t": "d", "a": "6181", "b": "74", "r": "c = 83, r = 39"}, {"t": "d", "a": "7505", "b": "73", "r": "c = 102, r = 59"}, {"t": "v", "a": "8,6", "s": "×", "b": "3,4", "r": "29,24"}]}, "2T": {"contenidos": "Fracciones: sumar, restar, multiplicar y dividir; división de decimales; raíz cuadrada.", "tiempo": 4, "mental": [{"q": "√49 =", "r": "7"}, {"q": "√4 =", "r": "2"}, {"q": "3/5 + 1/5 =", "r": "4/5"}, {"q": "2/10 + 1/10 =", "r": "3/10"}, {"q": "1/5 de 30 =", "r": "6"}, {"q": "9/10 de 20 =", "r": "18"}, {"q": "3/10 de 20 =", "r": "6"}, {"q": "El 25 % de 240 es", "r": "60"}, {"q": "3³ =", "r": "27"}, {"q": "4,9 × 1000 =", "r": "4900"}], "escrito": [{"t": "h", "q": "2/3 + 1/6 =", "r": "5/6"}, {"t": "h", "q": "4/5 − 4/10 =", "r": "2/5"}, {"t": "h", "q": "1/6 × 1/2 =", "r": "1/12"}, {"t": "h", "q": "1/3 : 1/2 =", "r": "2/3"}, {"t": "d", "a": "112,5", "b": "2,5", "r": "c = 45"}, {"t": "d", "a": "27,6", "b": "2,3", "r": "c = 12"}]}, "3T": {"contenidos": "Porcentajes; división entre decimales; potencias y operaciones combinadas; raíz cuadrada.", "tiempo": 4, "mental": [{"q": "El 20 % de 250 es", "r": "50"}, {"q": "El 50 % de 160 es", "r": "80"}, {"q": "El 50 % de 120 es", "r": "60"}, {"q": "El 10 % de 900 es", "r": "90"}, {"q": "√36 =", "r": "6"}, {"q": "2² =", "r": "4"}, {"q": "(2 + 4) × 4 =", "r": "24"}, {"q": "3/5 de 20 =", "r": "12"}, {"q": "87,1 : 100 =", "r": "0,871"}, {"q": "7,89 × 10 =", "r": "78,9"}], "escrito": [{"t": "h", "q": "El 40 % de 200 =", "r": "80"}, {"t": "h", "q": "El 20 % de 180 =", "r": "36"}, {"t": "d", "a": "91,2", "b": "1,2", "r": "c = 76"}, {"t": "d", "a": "43,8", "b": "0,6", "r": "c = 73"}, {"t": "h", "q": "2² + 2³ =", "r": "12"}, {"t": "h", "q": "60 − 7 × 8 =", "r": "4"}]}}};

// Banco de problemas: Parte 1 = un paso (3) · Parte 2 = dos pasos o razonamiento (3)
const PRUEBAS_RP = {"1º": {"1T": {"contenidos": "Problemas de juntar y quitar con números hasta 10, con apoyo de dibujo.", "p1": [{"e": "En la mesa hay 4 lápices rojos y 3 lápices azules. ¿Cuántos lápices hay en total?", "op": "4 + 3", "s": "Hay 7 lápices."}, {"e": "Lucía tiene 6 globos. Se le escapan 2. ¿Cuántos globos le quedan?", "op": "6 − 2", "s": "Le quedan 4 globos."}, {"e": "En el patio juegan 5 niños. Llegan 3 niños más. ¿Cuántos niños juegan ahora?", "op": "5 + 3", "s": "Juegan 8 niños."}], "p2": [{"e": "En un árbol hay 8 pájaros. Se van volando 3. ¿Cuántos pájaros quedan en el árbol?", "op": "8 − 3", "s": "Quedan 5 pájaros."}, {"e": "Pablo tiene 2 cromos. Su abuela le da 5 cromos. ¿Cuántos cromos tiene ahora Pablo?", "op": "2 + 5", "s": "Tiene 7 cromos."}, {"e": "En una caja hay 9 galletas. Nos comemos 4. ¿Cuántas galletas quedan?", "op": "9 − 4", "s": "Quedan 5 galletas."}]}, "2T": {"contenidos": "Problemas de juntar, añadir y quitar con números hasta 50.", "p1": [{"e": "En la clase de 1.º hay 12 niñas y 6 niños. ¿Cuántos alumnos hay en la clase?", "op": "12 + 6", "s": "Hay 18 alumnos."}, {"e": "Ana tenía 15 pegatinas. Ha regalado 4. ¿Cuántas pegatinas le quedan?", "op": "15 − 4", "s": "Le quedan 11 pegatinas."}, {"e": "En el autobús van 20 personas. Suben 7 más. ¿Cuántas personas van ahora?", "op": "20 + 7", "s": "Van 27 personas."}], "p2": [{"e": "Mario tiene 10 canicas y Sara tiene 3 más que Mario. ¿Cuántas canicas tiene Sara?", "op": "10 + 3", "s": "Sara tiene 13 canicas."}, {"e": "En el comedor había 28 yogures. Se han comido 5. ¿Cuántos yogures quedan?", "op": "28 − 5", "s": "Quedan 23 yogures."}, {"e": "En la biblioteca hay 30 cuentos en una estantería y 10 cuentos en otra. ¿Cuántos cuentos hay en total?", "op": "30 + 10", "s": "Hay 40 cuentos."}]}, "3T": {"contenidos": "Problemas de una operación (suma o resta) con números hasta 99; comparar.", "p1": [{"e": "En el huerto del cole hemos plantado 24 lechugas y 13 tomateras. ¿Cuántas plantas hemos plantado?", "op": "24 + 13", "s": "Hemos plantado 37 plantas."}, {"e": "Un libro tiene 58 páginas. Luis ha leído 25. ¿Cuántas páginas le faltan por leer?", "op": "58 − 25", "s": "Le faltan 33 páginas."}, {"e": "En una granja hay 41 gallinas y 26 patos. ¿Cuántas aves hay en total?", "op": "41 + 26", "s": "Hay 67 aves."}], "p2": [{"e": "Carmen tiene 35 euros y su hermano tiene 12 euros. ¿Cuántos euros tiene Carmen más que su hermano?", "op": "35 − 12", "s": "Carmen tiene 23 euros más."}, {"e": "En la excursión iban 47 alumnos. Se han quedado en el cole 5 alumnos de los 47. ¿Cuántos alumnos han ido a la excursión?", "op": "47 − 5", "s": "Han ido 42 alumnos."}, {"e": "Hugo tiene 2 decenas de cromos y 6 cromos sueltos. ¿Cuántos cromos tiene?", "op": "20 + 6", "s": "Tiene 26 cromos."}]}}, "2º": {"1T": {"contenidos": "Problemas de sumar y restar sin llevar con números hasta 99.", "p1": [{"e": "En el cole hay 34 alumnos en Infantil y 52 en el primer ciclo. ¿Cuántos alumnos hay en total?", "op": "34 + 52", "s": "Hay 86 alumnos."}, {"e": "Teníamos 68 pinturas y se han roto 25. ¿Cuántas pinturas nos quedan?", "op": "68 − 25", "s": "Nos quedan 43 pinturas."}, {"e": "Elena ha hecho 23 pulseras el sábado y 15 el domingo. ¿Cuántas pulseras ha hecho?", "op": "23 + 15", "s": "Ha hecho 38 pulseras."}], "p2": [{"e": "En un tren viajan 45 personas. En una estación bajan 12 y suben 6. ¿Cuántas personas van ahora en el tren?", "op": "45 − 12 = 33; 33 + 6", "s": "Van 39 personas."}, {"e": "Tengo 20 euros. Me compro un libro de 9 euros y un cuaderno de 1 euro. ¿Cuánto dinero me queda?", "op": "9 + 1 = 10; 20 − 10", "s": "Me quedan 10 euros."}, {"e": "Javi tiene 17 canicas, y su amiga Sofía, el doble. ¿Cuántas canicas tiene Sofía?", "op": "17 + 17", "s": "Sofía tiene 34 canicas."}]}, "2T": {"contenidos": "Problemas de sumar con llevadas y restar; primeras multiplicaciones (sumas repetidas).", "p1": [{"e": "En el mercadillo de Santa Fe venden 38 kilos de naranjas por la mañana y 27 kilos por la tarde. ¿Cuántos kilos venden en total?", "op": "38 + 27", "s": "Venden 65 kilos."}, {"e": "Un álbum tiene 150 cromos. Ya he pegado 40. ¿Cuántos cromos me faltan?", "op": "150 − 40", "s": "Me faltan 110 cromos."}, {"e": "Hay 5 cajas y en cada caja hay 2 pelotas. ¿Cuántas pelotas hay?", "op": "5 × 2", "s": "Hay 10 pelotas."}], "p2": [{"e": "Rocío tiene 3 bolsas con 10 caramelos cada una. Se come 4 caramelos. ¿Cuántos caramelos le quedan?", "op": "3 × 10 = 30; 30 − 4", "s": "Le quedan 26 caramelos."}, {"e": "En la clase de 2.º hay 16 alumnos y en la de 3.º, 9 más que en 2.º. ¿Cuántos alumnos hay en 3.º?", "op": "16 + 9", "s": "Hay 25 alumnos en 3.º."}, {"e": "Leo tenía 46 euros. Ha comprado un balón de 15 euros y le han regalado 10 euros. ¿Cuánto dinero tiene ahora?", "op": "46 − 15 = 31; 31 + 10", "s": "Tiene 41 euros."}]}, "3T": {"contenidos": "Problemas de sumar y restar con llevadas; multiplicar por una cifra.", "p1": [{"e": "Un coche tiene 4 ruedas. ¿Cuántas ruedas tienen 6 coches?", "op": "6 × 4", "s": "Tienen 24 ruedas."}, {"e": "En la biblioteca había 82 libros. Se han prestado 37. ¿Cuántos libros quedan?", "op": "82 − 37", "s": "Quedan 45 libros."}, {"e": "Un paquete de folios cuesta 3 euros. ¿Cuánto cuestan 5 paquetes?", "op": "5 × 3", "s": "Cuestan 15 euros."}], "p2": [{"e": "En una caja hay 2 filas de 10 bombones. Nos comemos 6. ¿Cuántos bombones quedan?", "op": "2 × 10 = 20; 20 − 6", "s": "Quedan 14 bombones."}, {"e": "Marta tiene 54 cromos y Pedro tiene 18 menos que Marta. ¿Cuántos cromos tienen entre los dos?", "op": "54 − 18 = 36; 54 + 36", "s": "Tienen 90 cromos."}, {"e": "Para la fiesta compramos 3 bolsas de 5 globos y 4 globos sueltos. ¿Cuántos globos tenemos?", "op": "3 × 5 = 15; 15 + 4", "s": "Tenemos 19 globos."}]}}, "3º": {"1T": {"contenidos": "Problemas de sumar y restar con llevadas (hasta 999); multiplicar por una cifra.", "p1": [{"e": "Un colegio tiene 168 alumnos y otro tiene 245. ¿Cuántos alumnos hay entre los dos colegios?", "op": "168 + 245", "s": "Hay 413 alumnos."}, {"e": "Un granjero tenía 520 huevos y ha vendido 285. ¿Cuántos huevos le quedan?", "op": "520 − 285", "s": "Le quedan 235 huevos."}, {"e": "Cada semana leo 25 páginas. ¿Cuántas páginas leo en 4 semanas?", "op": "25 × 4", "s": "Leo 100 páginas."}], "p2": [{"e": "En el cine hay 300 butacas. Por la mañana se ocupan 126 y por la tarde, 148. ¿Cuántas butacas se ocupan en total? ¿Cuántas más se ocupan por la tarde?", "op": "126 + 148 = 274; 148 − 126 = 22", "s": "Se ocupan 274 butacas; por la tarde, 22 más."}, {"e": "Lola tiene 3 billetes de 10 euros y 2 monedas de 2 euros. ¿Cuánto dinero tiene?", "op": "3 × 10 = 30; 2 × 2 = 4; 30 + 4", "s": "Tiene 34 euros."}, {"e": "Un tren lleva 6 vagones con 35 viajeros cada uno. En la estación bajan 40 viajeros. ¿Cuántos viajeros quedan?", "op": "6 × 35 = 210; 210 − 40", "s": "Quedan 170 viajeros."}]}, "2T": {"contenidos": "Problemas de multiplicar por una cifra; problemas de dos operaciones.", "p1": [{"e": "Una caja de lápices trae 12 lápices. ¿Cuántos lápices hay en 8 cajas?", "op": "12 × 8", "s": "Hay 96 lápices."}, {"e": "Una bici cuesta 189 euros y un casco, 37 euros. ¿Cuánto cuestan las dos cosas?", "op": "189 + 37", "s": "Cuestan 226 euros."}, {"e": "En la Alhambra entran 425 visitantes por la mañana. Por la tarde entran 160 menos. ¿Cuántos visitantes entran por la tarde?", "op": "425 − 160", "s": "Entran 265 visitantes."}], "p2": [{"e": "Compro 3 camisetas de 15 euros. Pago con un billete de 50 euros. ¿Cuánto me devuelven?", "op": "3 × 15 = 45; 50 − 45", "s": "Me devuelven 5 euros."}, {"e": "Un panadero hace 125 barras de pan cada día. ¿Cuántas barras hace de lunes a sábado (6 días)?", "op": "125 × 6", "s": "Hace 750 barras."}, {"e": "En el comedor hay 7 mesas con 8 sillas cada una. Ya están sentados 39 alumnos. ¿Cuántas sillas quedan libres?", "op": "7 × 8 = 56; 56 − 39", "s": "Quedan 17 sillas libres."}]}, "3T": {"contenidos": "Problemas de multiplicar y dividir por una cifra; repartos.", "p1": [{"e": "Repartimos 48 caramelos entre 6 niños a partes iguales. ¿Cuántos caramelos le tocan a cada uno?", "op": "48 : 6", "s": "A cada uno le tocan 8 caramelos."}, {"e": "Cada caja tiene 24 huevos. ¿Cuántos huevos hay en 5 cajas?", "op": "24 × 5", "s": "Hay 120 huevos."}, {"e": "Tenemos 63 fotos y las queremos poner en un álbum de 7 fotos por página. ¿Cuántas páginas llenaremos?", "op": "63 : 7", "s": "Llenaremos 9 páginas."}], "p2": [{"e": "La maestra tiene 90 cuadernos y los reparte entre 3 clases. Cada clase gasta 12. ¿Cuántos cuadernos le quedan a cada clase?", "op": "90 : 3 = 30; 30 − 12", "s": "A cada clase le quedan 18 cuadernos."}, {"e": "Juan ahorra 4 euros cada semana. ¿Cuánto ahorra en 9 semanas? Si quiere un juego de 45 euros, ¿le falta dinero?", "op": "4 × 9 = 36; 45 − 36 = 9", "s": "Ahorra 36 euros; le faltan 9 euros."}, {"e": "En una excursión van 85 alumnos en microbuses de 9 plazas. ¿Cuántos microbuses se llenan? ¿Cuántos alumnos sobran?", "op": "85 : 9 → cociente 9, resto 4", "s": "Se llenan 9 microbuses y sobran 4 alumnos (hace falta otro más)."}]}}, "4º": {"1T": {"contenidos": "Problemas de multiplicar por dos cifras y dividir por una cifra; dos operaciones.", "p1": [{"e": "Un autobús lleva 54 pasajeros. ¿Cuántos pasajeros llevan 12 autobuses iguales?", "op": "54 × 12", "s": "Llevan 648 pasajeros."}, {"e": "Repartimos 245 folios entre 5 clases por igual. ¿Cuántos folios recibe cada clase?", "op": "245 : 5", "s": "Cada clase recibe 49 folios."}, {"e": "El año pasado el colegio recogió 1.250 kilos de papel y este año 875 kilos. ¿Cuántos kilos recogió en los dos años?", "op": "1250 + 875", "s": "Recogió 2.125 kilos."}], "p2": [{"e": "Una entrada al parque cuesta 18 euros. Van 25 alumnos y pagan con 500 euros. ¿Cuánto dinero les devuelven?", "op": "18 × 25 = 450; 500 − 450", "s": "Les devuelven 50 euros."}, {"e": "Un agricultor recoge 728 kilos de tomates y los mete en cajas de 8 kilos. Vende cada caja a 6 euros. ¿Cuánto dinero obtiene?", "op": "728 : 8 = 91; 91 × 6", "s": "Obtiene 546 euros."}, {"e": "Una tienda tenía 1.500 botellas de agua. Vende 345 el lunes y 418 el martes. ¿Cuántas botellas le quedan?", "op": "345 + 418 = 763; 1500 − 763", "s": "Le quedan 737 botellas."}]}, "2T": {"contenidos": "Problemas de dividir por una y dos cifras; problemas de dos operaciones.", "p1": [{"e": "Un libro tiene 168 páginas. Si leo 12 páginas al día, ¿cuántos días tardo en leerlo?", "op": "168 : 12", "s": "Tardo 14 días."}, {"e": "Una fábrica hace 2.436 tornillos y los guarda en 6 cajas iguales. ¿Cuántos tornillos hay en cada caja?", "op": "2436 : 6", "s": "Hay 406 tornillos en cada caja."}, {"e": "Un televisor cuesta 389 euros. ¿Cuánto cuestan 15 televisores iguales?", "op": "389 × 15", "s": "Cuestan 5.835 euros."}], "p2": [{"e": "En una carrera solidaria se apuntan 264 alumnos. Forman equipos de 24. Cada equipo recibe 3 balones. ¿Cuántos balones hacen falta?", "op": "264 : 24 = 11; 11 × 3", "s": "Hacen falta 33 balones."}, {"e": "María compra 4 cuadernos de 3 euros y 2 estuches de 7 euros. Paga con 50 euros. ¿Cuánto le devuelven?", "op": "4 × 3 = 12; 2 × 7 = 14; 12 + 14 = 26; 50 − 26", "s": "Le devuelven 24 euros."}, {"e": "Tres amigos ganan 486 euros vendiendo camisetas. Se gastan 36 euros en la tela y reparten el resto a partes iguales. ¿Cuánto recibe cada uno?", "op": "486 − 36 = 450; 450 : 3", "s": "Cada uno recibe 150 euros."}]}, "3T": {"contenidos": "Problemas con dinero y decimales; división entre dos cifras; dos operaciones.", "p1": [{"e": "Un bocadillo cuesta 2,50 € y un zumo 1,25 €. ¿Cuánto cuestan las dos cosas?", "op": "2,50 + 1,25", "s": "Cuestan 3,75 €."}, {"e": "Tengo 10 € y gasto 6,40 € en una revista. ¿Cuánto dinero me queda?", "op": "10 − 6,40", "s": "Me quedan 3,60 €."}, {"e": "Hay 375 alumnos para ir de excursión en autobuses de 25 plazas. ¿Cuántos autobuses hacen falta?", "op": "375 : 25", "s": "Hacen falta 15 autobuses."}], "p2": [{"e": "Compro 3 bolígrafos de 1,20 € cada uno. Pago con un billete de 5 €. ¿Cuánto me devuelven?", "op": "3 × 1,20 = 3,60; 5 − 3,60", "s": "Me devuelven 1,40 €."}, {"e": "Un jardinero tiene 448 plantas y las coloca en 32 filas iguales. ¿Cuántas plantas pone en cada fila? Si cada planta cuesta 2 €, ¿cuánto vale una fila?", "op": "448 : 32 = 14; 14 × 2", "s": "Pone 14 plantas por fila; una fila vale 28 €."}, {"e": "Para el desayuno saludable compramos 4 kg de fruta a 2,25 € el kilo y pan por 3,50 €. ¿Cuánto gastamos?", "op": "4 × 2,25 = 9; 9 + 3,50", "s": "Gastamos 12,50 €."}]}}, "5º": {"1T": {"contenidos": "Problemas de varias operaciones con números naturales; división entre dos cifras.", "p1": [{"e": "Un camión transporta 1.536 kilos de naranjas en cajas de 24 kilos. ¿Cuántas cajas lleva?", "op": "1536 : 24", "s": "Lleva 64 cajas."}, {"e": "Una piscina municipal recibe 345 bañistas cada día. ¿Cuántos bañistas recibe en 31 días?", "op": "345 × 31", "s": "Recibe 10.695 bañistas."}, {"e": "Santa Fe tiene unos 15.000 habitantes y un pueblo vecino tiene 8.750. ¿Cuántos habitantes más tiene Santa Fe?", "op": "15000 − 8750", "s": "Tiene 6.250 habitantes más."}], "p2": [{"e": "Un colegio compra 12 ordenadores a 425 € cada uno. Recibe una ayuda de 2.000 €. ¿Cuánto tiene que pagar el colegio?", "op": "12 × 425 = 5.100; 5.100 − 2.000", "s": "Tiene que pagar 3.100 €."}, {"e": "En una fiesta hay 156 personas. Se sientan en mesas de 12. Cada mesa necesita 2 jarras de agua. ¿Cuántas jarras hacen falta?", "op": "156 : 12 = 13; 13 × 2", "s": "Hacen falta 26 jarras."}, {"e": "Laura ahorra 35 € al mes durante un año. Se compra una tablet de 299 €. ¿Cuánto dinero le sobra?", "op": "35 × 12 = 420; 420 − 299", "s": "Le sobran 121 €."}]}, "2T": {"contenidos": "Problemas con números decimales y fracciones de una cantidad.", "p1": [{"e": "Un kilo de manzanas cuesta 1,85 €. ¿Cuánto cuestan 4 kilos?", "op": "1,85 × 4", "s": "Cuestan 7,40 €."}, {"e": "Una cuerda mide 12,5 m. Cortamos 3,75 m. ¿Cuántos metros quedan?", "op": "12,5 − 3,75", "s": "Quedan 8,75 m."}, {"e": "En una clase de 24 alumnos, 3/4 van al comedor. ¿Cuántos alumnos van al comedor?", "op": "24 : 4 = 6; 6 × 3", "s": "Van 18 alumnos."}], "p2": [{"e": "Cinco amigos pagan a partes iguales una pizza de 18,50 € y bebidas por 6,50 €. ¿Cuánto paga cada uno?", "op": "18,50 + 6,50 = 25; 25 : 5", "s": "Cada uno paga 5 €."}, {"e": "Tengo 60 €. Gasto 1/3 en un libro y 1/4 en una camiseta. ¿Cuánto dinero me queda?", "op": "60 : 3 = 20; 60 : 4 = 15; 60 − 20 − 15", "s": "Me quedan 25 €."}, {"e": "Una botella tiene 1,5 litros de agua. ¿Cuántos litros hay en un pack de 6 botellas? ¿Y en 2 packs?", "op": "1,5 × 6 = 9; 9 × 2", "s": "En un pack hay 9 litros y en dos, 18 litros."}]}, "3T": {"contenidos": "Problemas con porcentajes sencillos, fracciones y decimales; varias operaciones.", "p1": [{"e": "Unas zapatillas cuestan 60 € y tienen un 10 % de descuento. ¿Cuántos euros nos descuentan?", "op": "10 % de 60 = 60 : 10", "s": "Nos descuentan 6 €."}, {"e": "De los 200 alumnos de un colegio, el 50 % son chicas. ¿Cuántas chicas hay?", "op": "200 : 2", "s": "Hay 100 chicas."}, {"e": "Un ciclista recorre 7,8 km cada día. ¿Cuántos kilómetros recorre en 5 días?", "op": "7,8 × 5", "s": "Recorre 39 km."}], "p2": [{"e": "Un abrigo cuesta 80 € y lo rebajan un 25 %. ¿Cuánto cuesta ahora?", "op": "25 % de 80 = 20; 80 − 20", "s": "Cuesta 60 €."}, {"e": "Un depósito tiene 450 litros. Se gastan 2/5 del agua. ¿Cuántos litros quedan?", "op": "450 : 5 × 2 = 180; 450 − 180", "s": "Quedan 270 litros."}, {"e": "Compramos 3 kg de queso a 9,60 € el kilo y lo repartimos en 4 bolsas iguales. ¿Cuánto cuesta el queso de cada bolsa?", "op": "3 × 9,60 = 28,80; 28,80 : 4", "s": "Cada bolsa cuesta 7,20 €."}]}}, "6º": {"1T": {"contenidos": "Problemas de varias operaciones con naturales y decimales; potencias en contextos sencillos.", "p1": [{"e": "Un instituto encarga 24 cajas con 125 libros cada una. ¿Cuántos libros recibe?", "op": "24 × 125", "s": "Recibe 3.000 libros."}, {"e": "Un depósito de gasoil de 1.000 litros tiene 357,5 litros. ¿Cuántos litros faltan para llenarlo?", "op": "1000 − 357,5", "s": "Faltan 642,5 litros."}, {"e": "Un jardín cuadrado mide 9 m de lado. ¿Cuál es su área?", "op": "9² = 9 × 9", "s": "Su área es 81 m²."}], "p2": [{"e": "Una familia gasta 85,40 € a la semana en comida. ¿Cuánto gasta en 4 semanas? Si gana 1.500 € al mes, ¿cuánto le queda para otros gastos?", "op": "85,40 × 4 = 341,60; 1.500 − 341,60", "s": "Gasta 341,60 € y le quedan 1.158,40 €."}, {"e": "Un móvil cuesta 360 €. Se paga 60 € al principio y el resto en 12 plazos iguales. ¿Cuánto es cada plazo?", "op": "360 − 60 = 300; 300 : 12", "s": "Cada plazo es de 25 €."}, {"e": "En un almacén hay 8 estanterías con 8 cajas cada una y cada caja tiene 8 botes. ¿Cuántos botes hay? (Exprésalo como potencia).", "op": "8 × 8 × 8 = 8³", "s": "Hay 512 botes (8³)."}]}, "2T": {"contenidos": "Problemas con fracciones, decimales y divisiones con decimales.", "p1": [{"e": "En un viaje de 240 km ya hemos recorrido 5/8. ¿Cuántos kilómetros hemos recorrido?", "op": "240 : 8 × 5", "s": "Hemos recorrido 150 km."}, {"e": "Con 22,5 m de tela se hacen cortinas de 2,5 m cada una. ¿Cuántas cortinas salen?", "op": "22,5 : 2,5", "s": "Salen 9 cortinas."}, {"e": "Pedro se come 1/4 de una pizza y Ana 2/4. ¿Qué fracción de pizza se han comido entre los dos?", "op": "1/4 + 2/4", "s": "Se han comido 3/4 de la pizza."}], "p2": [{"e": "De una tarta, Luis se come 1/3 y su hermana 1/6. ¿Qué fracción de tarta queda?", "op": "1/3 + 1/6 = 2/6 + 1/6 = 3/6 = 1/2; 1 − 1/2", "s": "Queda 1/2 de la tarta."}, {"e": "Un depósito de 1.200 litros está lleno en sus 3/4 partes. Se gastan 350 litros. ¿Cuántos litros quedan?", "op": "1.200 : 4 × 3 = 900; 900 − 350", "s": "Quedan 550 litros."}, {"e": "Un kilo de almendras cuesta 12,80 €. ¿Cuánto cuestan 0,75 kg? ¿Cuánto me devuelven si pago con 20 €?", "op": "12,80 × 0,75 = 9,60; 20 − 9,60", "s": "Cuestan 9,60 € y me devuelven 10,40 €."}]}, "3T": {"contenidos": "Problemas con porcentajes, fracciones y decimales; varias operaciones.", "p1": [{"e": "En un colegio de 250 alumnos, el 20 % usa el transporte escolar. ¿Cuántos alumnos lo usan?", "op": "250 × 20 : 100", "s": "Lo usan 50 alumnos."}, {"e": "Un ordenador cuesta 640 € y sube un 15 %. ¿Cuántos euros sube?", "op": "640 × 15 : 100", "s": "Sube 96 €."}, {"e": "Un coche gasta 6,5 litros cada 100 km. ¿Cuántos litros gasta en 400 km?", "op": "6,5 × 4", "s": "Gasta 26 litros."}], "p2": [{"e": "Una bicicleta cuesta 320 € y la rebajan un 30 %. ¿Cuál es el precio final?", "op": "30 % de 320 = 96; 320 − 96", "s": "El precio final es 224 €."}, {"e": "Una clase recauda 450 € en un mercadillo. Dedica el 40 % a una ONG y el resto se reparte a partes iguales entre 3 actividades. ¿Cuánto recibe cada actividad?", "op": "40 % de 450 = 180; 450 − 180 = 270; 270 : 3", "s": "Cada actividad recibe 90 €."}, {"e": "Un granjero tiene 1.800 kg de aceitunas. Vende 2/3 a 1,50 € el kilo. ¿Cuánto dinero obtiene?", "op": "1.800 : 3 × 2 = 1.200; 1.200 × 1,50", "s": "Obtiene 1.800 €."}]}}};

// ── librería docx (se carga solo al descargar) ─────────────────
const PW_DOCX_URL = 'docx-9.5.1.iife.js';   // librería docx 9.5.1 (MIT), copiada en el repositorio
let PW_DOCX_P = null;
function pwCargarDocx(){
  if(window.docx) return Promise.resolve(window.docx);
  if(!PW_DOCX_P) PW_DOCX_P = new Promise((ok, ko) => {
    const s = document.createElement('script');
    s.src = PW_DOCX_URL;
    s.onload = () => window.docx ? ok(window.docx) : ko(new Error('docx'));
    s.onerror = () => { PW_DOCX_P = null; ko(new Error('No se ha podido cargar el generador de Word. Comprueba la conexión.')); };
    document.head.appendChild(s);
  });
  return PW_DOCX_P;
}
let PW_LOGO = null;
async function pwLogo(){
  if(PW_LOGO !== null) return PW_LOGO;
  try{
    const r = await fetch('logo-cole.png', {cache:'force-cache'});
    if(!r.ok) throw 0;
    PW_LOGO = new Uint8Array(await r.arrayBuffer());
  }catch(e){ PW_LOGO = false; }
  return PW_LOGO;
}

// ── utilidades de maquetación ──────────────────────────────────
const CM = 567;                 // twips por centímetro
const PW_W = 9864;              // ancho útil (A4 − 2 × 1,8 cm)
let D = null;                   // espacio de nombres de docx, cuando está cargado
const hex = c => (c || '1a2340').replace('#', '');

// texto con **negrita**
function pwRuns(txt, o = {}){
  if(Array.isArray(txt)) return txt;
  const out = [];
  String(txt).split(/(\*\*[^*]+\*\*)/).forEach(seg => {
    if(!seg) return;
    const b = seg.startsWith('**') && seg.endsWith('**');
    out.push(new D.TextRun({ text: b ? seg.slice(2, -2) : seg, bold: b || o.b, italics: o.i,
      size: o.sz ? Math.round(o.sz * 2) : undefined, color: o.c ? hex(o.c) : undefined, font: o.f }));
  });
  return out;
}
function P(txt, o = {}){
  const al = { c: D.AlignmentType.CENTER, r: D.AlignmentType.RIGHT }[o.al];
  return new D.Paragraph({ children: pwRuns(txt, o), alignment: al,
    spacing: { before: o.before || 0, after: o.after === undefined ? 100 : o.after, line: o.line },
    indent: o.ind ? { left: o.ind, hanging: o.hang || 0 } : undefined,
    border: o.linea ? { bottom: { style: D.BorderStyle.DOTTED, size: 6, color: '999999', space: 1 } } : undefined,
    keepNext: o.keepNext });
}
const H = (txt, col, sz = 13, o = {}) => P(txt, Object.assign({ b: true, sz, c: col, before: o.before === undefined ? 120 : o.before, after: 80, keepNext: true }, o));
const lista = (items, sz = 11) => items.map((t, i) => P(`${i + 1}.  ${t}`, { sz, ind: 360, hang: 360, after: 60 }));
const NB = () => ({ style: D.BorderStyle.NONE, size: 0, color: 'FFFFFF' });
const LB = (c = 'BBBBBB', s = 6) => ({ style: D.BorderStyle.SINGLE, size: s, color: hex(c) });
function bordes(tipo, c, s){
  if(tipo === 'none') return { top: NB(), bottom: NB(), left: NB(), right: NB() };
  if(typeof tipo === 'object') return tipo;
  return { top: LB(c, s), bottom: LB(c, s), left: LB(c, s), right: LB(c, s) };
}
// celda: contenido = texto | Paragraph | Table | array
function C(cont, o = {}){
  let ch = Array.isArray(cont) ? cont : [cont];
  ch = ch.map(x => (typeof x === 'string' || typeof x === 'number') ? P(String(x), { sz: o.sz || 10.5, b: o.b, al: o.al, c: o.c, after: 0 }) : x);
  if(!ch.length || !(ch[ch.length - 1] instanceof D.Paragraph)) ch.push(new D.Paragraph({ children: [], spacing: { after: 0 } }));
  return new D.TableCell({ children: ch, width: o.w ? { size: o.w, type: D.WidthType.DXA } : undefined,
    columnSpan: o.span, verticalAlign: o.va === 'c' ? D.VerticalAlign.CENTER : o.va === 'b' ? D.VerticalAlign.BOTTOM : D.VerticalAlign.TOP,
    shading: o.bg ? { type: D.ShadingType.CLEAR, color: 'auto', fill: hex(o.bg) } : undefined,
    margins: { top: o.pad === undefined ? 70 : o.pad, bottom: o.pad === undefined ? 70 : o.pad, left: 100, right: 100 },
    borders: o.borders });
}
function T(filas, o = {}){
  const anchos = o.anchos || [];
  const total = o.ancho || PW_W;
  return new D.Table({
    width: { size: total, type: D.WidthType.DXA },
    columnWidths: anchos.length ? anchos : undefined,
    layout: D.TableLayoutType.FIXED,
    alignment: o.centro ? D.AlignmentType.CENTER : undefined,
    borders: bordes(o.bordes === undefined ? 'all' : o.bordes, o.color, o.grosor),
    rows: filas.map((f, ri) => new D.TableRow({
      children: f.map((c, ci) => c instanceof D.TableCell ? c : C(c, Object.assign({ w: anchos[ci], bg: (o.cabecera && ri === 0) ? o.cabecera : undefined, b: o.cabecera && ri === 0 }, o.celda || {}))),
      height: (o.altos && o.altos[ri]) ? { value: o.altos[ri], rule: D.HeightRule.ATLEAST } : o.alto ? { value: o.alto, rule: D.HeightRule.ATLEAST } : undefined,
      cantSplit: true,
      tableHeader: o.cabecera && ri === 0,
    })),
  });
}
const anchosPct = pcts => { const a = pcts.map(p => Math.round(PW_W * p / 100)); return a; };

function pwCabecera(izq, der, col){
  const celdas = [];
  const bb = { top: NB(), left: NB(), right: NB(), bottom: LB(col, 16) };
  if(PW_LOGO){
    celdas.push(C(new D.Paragraph({ children: [new D.ImageRun({ type: 'png', data: PW_LOGO, transformation: { width: 50, height: 43 } })], spacing: { after: 0 } }), { w: 1000, borders: bb, va: 'b', pad: 30 }));
  }
  const wl = PW_LOGO ? 5264 : 6264;
  celdas.push(C([P('CEIP CAPITULACIONES DE SANTA FE', { sz: 9, b: true, c: '#444444', after: 0 }), P(izq, { sz: 9, c: '#555555', after: 0 })], { w: wl, borders: bb, va: 'b', pad: 30 }));
  celdas.push(C(P(der, { sz: 9, c: '#555555', al: 'r', after: 0 }), { w: 3600, borders: bb, va: 'b', pad: 30 }));
  return new D.Header({ children: [T([celdas], { bordes: 'none', anchos: PW_LOGO ? [1000, 5264, 3600] : [6264, 3600] }), P('', { after: 0 })] });
}
// una hoja = una sección (salto de página + cabecera propia)
function hoja(cab, children){ return { cab, children }; }
async function pwDescargar(nombre, titulo, hojas){
  const doc = new D.Document({
    creator: 'CEIP Capitulaciones de Santa Fe', title: titulo,
    styles: { default: { document: { run: { font: 'Calibri', size: 22, color: '1A2340' } } } },
    sections: hojas.map(h => ({
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1300, bottom: 900, left: 1021, right: 1021, header: 420, footer: 400 } } },
      headers: { default: pwCabecera(h.cab[0], h.cab[1], h.cab[2]) },
      children: h.children,
    })),
  });
  const blob = await D.Packer.toBlob(doc);
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = nombre;
  document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
}
async function pwPreparar(){ D = await pwCargarDocx(); await pwLogo(); }

function pwMomentos(keys){ return MOMENTOS.filter(m => keys.includes(m.key)); }
function pwNombreTrim(keys){ return keys.length === 3 ? 'curso-completo' : keys[0]; }
function pwDatosAlumno(){
  return [T([[C(P('Nombre:', { sz: 11, after: 0 }), { w: 6864, pad: 120 }), C(P('Fecha:  ___ / ___ / _____', { sz: 11, after: 0 }), { w: 3000, pad: 120 })]], { anchos: [6864, 3000] }), P('', { after: 120 })];
}
function pwTitulo(t, sub){ return [P(t, { b: true, sz: 17, al: 'c', after: 0 }), P(sub, { sz: 10.5, c: '#777777', al: 'c', after: 80 })]; }
function pwRegistro(cursoId, m, col, cabTxt, cols, extra){
  const c = S.cursos[cursoId];
  const als = (c && c.alumnos) ? c.alumnos : [];
  const nombres = siglasUnicas(als.map(a => a.nombre));
  const lst = nombres.length ? nombres : new Array(14).fill('');
  const wNom = 2200, wc = Math.floor((PW_W - wNom) / cols.length);
  const anchos = [wNom].concat(cols.map(() => wc));
  const filas = [['Alumno/a'].concat(cols)].concat(lst.map(n => [n].concat(cols.map(() => ''))));
  return hoja([cabTxt, `${cursoId} · ${m.corto} · ${anioLabel(S.anio)}`, col], [
    H(`Hoja de registro del grupo · ${cursoId} · ${m.corto}`, col, 13, { before: 0 }),
    P(`Fecha de aplicación: ___ / ___ / _____      Maestro/a: ____________________________${extra ? '      ' + extra : ''}`, { sz: 10, c: '#555555', after: 120 }),
    T(filas, { anchos, cabecera: '#EEF3FF', celda: { sz: 10 }, alto: 360 }),
    P('Documento con datos del alumnado: archívese en el expediente del Plan de Lectura y Matemáticas.', { sz: 9.5, c: '#777777', before: 120 }),
  ]);
}

// ══════════════════════════════════════════════════════════════
// LECTURA
// ══════════════════════════════════════════════════════════════
async function generarWord(cursoId, keys, tipo){
  await pwPreparar();
  const base = cursoBase(cursoId);
  if(!base || !PRUEBAS[base]){ alert('No hay prueba definida para este curso.'); return; }
  keys = keys || TS; tipo = tipo || 'alumno';
  const peque = (base === '1º' || base === '2º');
  const fz = peque ? 19 : (base === '3º' || base === '4º') ? 15 : 13;
  const lh = peque ? 552 : (base === '3º' || base === '4º') ? 456 : 410;
  const COL = '#4f7ef8';
  const cab = m => ['Plan de Lectura · Prueba de lectura', `${cursoId} · ${m.corto} · ${anioLabel(S.anio)}`, COL];
  const hojas = [];

  if(tipo === 'maestro'){
    hojas.push(hoja(['Plan de Lectura · Cuadernillo del maestro/a', `${cursoId} · ${anioLabel(S.anio)}`, COL], [
      H(`Pruebas de lectura · ${cursoId}`, COL, 18, { before: 0 }),
      P(`Curso escolar ${anioLabel(S.anio)} · Código de centro 18008853`, { sz: 11, c: '#555555' }),
      P('Este es el **cuadernillo del maestro/a**. Las fichas para el alumnado (texto y preguntas) van en otro documento, el **cuadernillo del alumnado**, para fotocopiarlo sin las soluciones.', { sz: 11, after: 160 }),
      T([['Hoja', 'Contenido'],
         ['**Hoja de lectura**', 'El texto con el recuento de palabras y las casillas de errores y ppm'],
         ['**Guía (2 hojas)**', 'Aplicación, baremo de velocidad, respuestas orientativas, rúbrica de expresión oral y volcado en la app'],
         ['**Hoja de registro**', 'Tabla del grupo con la lista de clase']], { anchos: [3000, 6864], cabecera: '#EEF3FF', celda: { sz: 11 } }),
      P('Textos incluidos:', { sz: 11, before: 200 }),
      T([['Trimestre', 'Texto', 'Palabras', 'Objetivo']].concat(pwMomentos(keys).map(m => {
        const pr = PRUEBAS[base][m.key];
        return [`${m.corto} (${m.epoca})`, `«${pr.titulo}»`, String(palabras(pr.texto)), `${VLOBJ[base][m.key]} ppm`];
      })), { anchos: [3000, 3864, 1500, 1500], cabecera: '#EEF3FF', celda: { sz: 11 } }),
    ]));
  }

  pwMomentos(keys).forEach(m => {
    const pr = PRUEBAS[base][m.key];
    const obj = VLOBJ[base][m.key];
    const np = palabras(pr.texto);
    const tr = tramosVL(obj);

    if(tipo === 'alumno'){
      const parrafos = pr.texto.split('\n').filter(l => l.trim()).map(l => P(l.trim(), { sz: fz, line: lh, after: 20 }));
      hojas.push(hoja(cab(m), [...pwTitulo(pr.titulo, 'Lectura en voz alta'), ...pwDatosAlumno(),
        T([[C(parrafos, { pad: 200 })]], { anchos: [PW_W], color: 'DDDDDD' })]));
      const szq = peque ? 14 : 12;
      const preg = [];
      const hl = peque ? 520 : 360, nq = pr.preguntas.length;
      const nl = peque ? 2 : Math.max(2, Math.min(3, Math.floor((11400 / nq - 330) / hl)));
      pr.preguntas.forEach((q, i) => {
        preg.push(P(`**${i + 1}.**  ${q.p}`, { sz: szq, before: 80, after: 0, keepNext: true }));
        const dot = { top: NB(), left: NB(), right: NB(), bottom: { style: D.BorderStyle.DOTTED, size: 6, color: '999999' } };
        preg.push(T(new Array(nl).fill(0).map(() => [C(P('', { after: 0, sz: 8 }), { w: PW_W, borders: dot, pad: 0 })]), { anchos: [PW_W], bordes: 'none', alto: hl }));
      });
      hojas.push(hoja(cab(m), [...pwTitulo(pr.titulo, 'Comprensión lectora'), ...pwDatosAlumno(),
        P(peque ? 'Escucha bien la pregunta y responde.' : 'Puedes volver a leer el texto si lo necesitas. Responde con frases completas.', { sz: 10.5, c: '#666666', i: true }),
        ...preg]));
      return;
    }

    // ── maestro/a ──
    const lin = lineasContadas(pr.texto);
    hojas.push(hoja(cab(m), [
      H('Hoja de lectura · copia del maestro/a', COL, 13, { before: 0 }),
      P(`«${pr.titulo}» · ${np} palabras · objetivo del nivel: **${obj} ppm**. Marca con una barra ( / ) la palabra en la que se encuentra al cumplirse el minuto, pero deja que termine el texto. La columna de la derecha da las palabras acumuladas al final de cada línea.`, { sz: 10, c: '#555555' }),
      P('Alumno/a: _____________________________________   Fecha: ___ / ___ / _____', { sz: 10, c: '#555555', after: 120 }),
      T(lin.map(l => [C(P(l.t, { sz: peque ? 15 : 12, after: 0, line: peque ? 380 : (base === '5º' || base === '6º') ? 290 : 340 }), { w: 8964, borders: { top: NB(), bottom: NB(), left: NB(), right: { style: D.BorderStyle.DOTTED, size: 4, color: 'BBBBBB' } } }),
                      C(P(String(l.n), { sz: 9, c: '#999999', al: 'r', after: 0 }), { w: 900, va: 'c', borders: bordes('none') })]),
        { anchos: [8964, 900], color: 'DDDDDD' }),
      P('', { after: 120 }),
      T([['Palabras en 1 minuto', 'Errores', 'PPM correctas'], ['', '', '']], { anchos: [3288, 3288, 3288], celda: { sz: 11 }, alto: 500 }),
      P('Errores: sustituciones, omisiones, adiciones, inversiones y palabras dictadas. Las autocorrecciones inmediatas no cuentan.', { sz: 10, c: '#555555', before: 100 }),
    ]));

    hojas.push(hoja(cab(m), [
      H('Guía del maestro/a (1 de 2) · aplicación y velocidad lectora', COL, 13, { before: 0 }),
      H('1 · Cómo se aplica', COL, 12),
      ...lista([
        '**Individualmente**, fuera del aula o en un rincón tranquilo. Entre 4 y 6 minutos por alumno/a.',
        'Entrega la **ficha del texto** (cuadernillo del alumnado). Tú usas la **hoja de lectura** para cronometrar y anotar.',
        'Consigna literal: «Vas a leer este texto en voz alta. Léelo como tú sabes, sin correr. Cuando termines te haré unas preguntas.» No se le dice que se le cronometra, para no inducir prisa.',
        'Cronómetro al empezar la primera palabra. Al minuto exacto marcas una barra, pero **le dejas terminar**.',
        'Retira el texto y entrega la **ficha de preguntas**.' + (peque ? ' En 1º y 2º las preguntas se hacen **oralmente** y las respuestas las anotas tú en la ficha.' : ''),
        'Termina pidiéndole que cuente el texto con sus palabras: eso es lo que valoras en expresión oral.']),
      H('2 · Velocidad lectora (VL)', COL, 12),
      P('**PPM correctas = palabras leídas en 1 minuto − errores.**', { sz: 11 }),
      T([['PPM correctas', 'Valoración', 'Nota (0–10)']].concat(tr.map(x => [`${x.min}${x.max === null ? ' o más' : ' – ' + x.max} ppm`, x.lbl, `**${String(x.nota).replace('.', ',')}**`])),
        { anchos: [3288, 4288, 2288], cabecera: '#EEF3FF', celda: { sz: 11 } }),
      P(`Baremo del centro para ${base}, ${m.corto}, con **${obj} ppm** como objetivo. Los objetivos de 1º y 2º están ajustados a la baja respecto a las tablas generales, teniendo en cuenta el contexto del centro (Zona de Transformación Social). Puedes matizar la nota medio punto según entonación, respeto de las pausas y fluidez.`, { sz: 10, c: '#555555', before: 100 }),
    ]));

    hojas.push(hoja(cab(m), [
      H('Guía del maestro/a (2 de 2) · comprensión, expresión oral y volcado', COL, 13, { before: 0 }),
      H('3 · Comprensión lectora (CL)', COL, 12),
      P(`${pr.preguntas.length} preguntas × **${(10 / pr.preguntas.length).toFixed(2).replace('.', ',')} puntos**: completa (todo), parcial (la mitad), incorrecta o en blanco (0). Las preguntas Personal o Crítica se puntúan por la coherencia del razonamiento, no por la opinión. **Nota CL = (puntos ÷ ${pr.preguntas.length}) × 10.**`, { sz: 11 }),
      T([['Nº', 'Tipo', 'Respuesta orientativa']].concat(pr.preguntas.map((q, i) => [String(i + 1), `**${q.t}**`, q.r])),
        { anchos: [500, 1500, 7864], cabecera: '#EEF3FF', celda: { sz: 10 } }),
      H('4 · Expresión oral (EO)', COL, 12),
      P('«Cuéntame de qué trata el texto, con tus palabras». Cuatro aspectos de 2,5 puntos.', { sz: 11 }),
      T([['Aspecto', '0 – 1', '1,5', '2 – 2,5'],
         ['**Contenido**', 'Datos sueltos o erróneos', 'Idea general sin detalles', 'Idea principal y detalles'],
         ['**Orden**', 'Sin secuencia clara', 'Orden con algún salto', 'Principio, desarrollo y final'],
         ['**Vocabulario**', 'Muy pobre o repetitivo', 'Correcto y sencillo', 'Preciso, usa palabras del texto'],
         ['**Fluidez**', 'Muchas pausas y muletillas', 'Alguna vacilación', 'Continuo y bien entonado']],
        { anchos: [1800, 2688, 2688, 2688], cabecera: '#EEF3FF', celda: { sz: 10 } }),
      H('5 · Cómo se meten los datos en la app', COL, 12),
      ...lista([
        `Entra en **Registro**, elige el curso escolar **${anioLabel(S.anio)}**, el curso **${cursoId}** y el **${m.corto}**.`,
        'Pulsa sobre las siglas del alumno/a para abrir su ficha.',
        'Escribe la nota de **Velocidad Lectora**, la de **Comprensión Lectora** y la de **Expresión Oral**. Admite decimales. Se guarda solo.']),
      P('**Importante:** en la app se guarda la nota de 0 a 10, no las ppm. Las ppm quedan en la hoja de registro de papel.', { sz: 10, c: '#555555' }),
    ]));

    hojas.push(pwRegistro(cursoId, m, COL, 'Plan de Lectura · Prueba de lectura',
      ['Palabras 1 min', 'Errores', 'PPM', 'Nota VL', 'Nota CL', 'Nota EO'], `Objetivo: ${obj} ppm`));
  });

  const quien = tipo === 'alumno' ? 'alumnado' : 'maestro';
  await pwDescargar(`Lectura_${cursoId.replace(/[^\w]/g, '')}_${pwNombreTrim(keys)}_${quien}_${S.anio}.docx`,
    `Pruebas de lectura ${cursoId} · ${quien}`, hojas);
}

// ══════════════════════════════════════════════════════════════
// CÁLCULO
// ══════════════════════════════════════════════════════════════
const PW_NUMF = 'Arial';
function pwOperacion(o, sz){
  const sp = s => s.split('').join(' ');      // separa las cifras para que se vean las columnas
  if(o.t === 'v'){
    const lb = { top: NB(), left: NB(), right: NB(), bottom: LB('1a2340', 14) };
    return T([
      [C(P('', { after: 0 }), { w: 500, borders: bordes('none'), pad: 0 }), C(P(sp(o.a), { sz, f: PW_NUMF, al: 'r', after: 0 }), { w: 1900, borders: bordes('none'), pad: 0 })],
      [C(P(o.s, { sz, f: PW_NUMF, after: 0 }), { w: 500, borders: lb, pad: 0 }), C(P(sp(o.b), { sz, f: PW_NUMF, al: 'r', after: 0 }), { w: 1900, borders: lb, pad: 0 })],
    ], { anchos: [500, 1900], bordes: 'none', ancho: 2400, centro: true });
  }
  if(o.t === 'd'){
    const db = { top: NB(), right: NB(), left: LB('1a2340', 14), bottom: LB('1a2340', 14) };
    return T([[C(P(sp(o.a), { sz, f: PW_NUMF, al: 'r', after: 0 }), { w: 1500, borders: bordes('none'), pad: 0 }),
               C(P(sp(o.b), { sz, f: PW_NUMF, after: 0 }), { w: 1000, borders: db, pad: 0 })]],
      { anchos: [1500, 1000], bordes: 'none', ancho: 2500, centro: true });
  }
  return P(o.q, { sz: sz - 2, f: PW_NUMF, al: 'c', after: 0 });
}
function pwNotaTabla(max, paso){
  const pts = [], notas = [];
  for(let p = 0; p <= max + 1e-9; p += paso){
    pts.push(String(p).replace('.', ','));
    notas.push('**' + String(Math.round(p / max * 1000) / 100).replace('.', ',') + '**');
  }
  const w0 = 1100, w = Math.floor((PW_W - w0) / pts.length);
  return T([['**Puntos**'].concat(pts), ['**Nota**'].concat(notas)], { anchos: [w0].concat(pts.map(() => w)), celda: { sz: 9, al: 'c' } });
}

async function generarCalculoWord(cursoId, keys, tipo){
  await pwPreparar();
  const base = cursoBase(cursoId);
  if(!base || !PRUEBAS_CALC[base]){ alert('No hay prueba de cálculo para este curso.'); return; }
  const peque = (base === '1º' || base === '2º');
  const COL = '#d97706';
  const cab = m => ['Plan de Matemáticas · Prueba de cálculo', `${cursoId} · ${m.corto} · ${anioLabel(S.anio)}`, COL];
  const hojas = [];
  pwMomentos(keys).forEach(m => {
    const pr = PRUEBAS_CALC[base][m.key];
    if(tipo === 'alumno'){
      const sz = peque ? 18 : 14;
      const caja = () => C(P('', { after: 0 }), { w: 1300, borders: bordes('all', '1a2340', 12) });
      const item = (it, i) => C(P(`**${i + 1}.**  ${it.q}`, { sz, f: PW_NUMF, after: 0 }), { w: 3432, borders: bordes('none'), va: 'c' });
      const hueco = () => C(P('', { after: 0 }), { w: 400, borders: bordes('none') });
      const filas = [];
      for(let i = 0; i < pr.mental.length; i += 2){
        filas.push([item(pr.mental[i], i), caja(), hueco(), item(pr.mental[i + 1], i + 1), caja()]);
      }
      hojas.push(hoja(cab(m), [...pwTitulo('Cálculo · Parte 1', 'Cálculo mental'), ...pwDatosAlumno(),
        P(peque ? 'Escucha y escribe el resultado en la casilla.' : 'Calcula de cabeza y escribe solo el resultado en la casilla. No hagas la cuenta en el papel.', { sz: peque ? 13 : 11, c: '#444444', after: 160 }),
        T(filas, { anchos: [3432, 1300, 400, 3432, 1300], bordes: 'none', alto: peque ? 1000 : 900 })]));

      const szo = peque ? 20 : 16;
      const alto = base === '1º' ? 2600 : base === '2º' ? 3200 : 4300;
      const celdaOp = (o, i) => C([P(String.fromCharCode(97 + i) + ')', { sz: 10, c: '#777777', after: 60 }), pwOperacion(o, szo)], { w: 3288, borders: bordes('all', 'CCCCCC') });
      const ops = pr.escrito;
      const instr = 'Haz las operaciones en su recuadro.' + (ops.some(o => o.t === 'h') ? ' Escribe los pasos y el resultado final.' : '') + (ops.some(o => o.t === 'd') ? ' En las divisiones, escribe el cociente y el resto.' : '');
      hojas.push(hoja(cab(m), [...pwTitulo('Cálculo · Parte 2', 'Operaciones'), ...pwDatosAlumno(),
        P(instr, { sz: peque ? 13 : 11, c: '#444444', after: 160 }),
        T([ops.slice(0, 3).map((o, i) => celdaOp(o, i)), ops.slice(3, 6).map((o, i) => celdaOp(o, i + 3))], { anchos: [3288, 3288, 3288], alto, color: 'CCCCCC' })]));
      return;
    }
    // ── maestro/a ──
    const hayDiv = pr.escrito.some(o => o.t === 'd');
    hojas.push(hoja(cab(m), [
      H(`Guía del maestro/a · Prueba de cálculo · ${base} · ${m.corto}`, COL, 13, { before: 0 }),
      P(`**Qué mide:** ${pr.contenidos}`, { sz: 10.5, c: '#555555' }),
      H('1 · Cómo se aplica', COL, 12),
      ...lista([
        'Colectiva, en el aula. Mejor en **dos sesiones** (un día cada parte) o con un descanso entre ellas.',
        '**Parte 1 · Cálculo mental**: ' + (peque ? `**lee tú en voz alta cada ítem dos veces** y deja unos 20–30 segundos para escribir el resultado (unos ${pr.tiempo + 3} minutos en total).` : `el alumnado la hace solo, sin hacer cuentas en el papel. Tiempo: **${pr.tiempo} minutos**.`),
        `**Parte 2 · Operaciones**: sin calculadora y sin límite estricto de tiempo (orientativo: ${peque ? '20' : '30'} minutos).` + (peque ? ' En 1º y 2º se puede leer en voz alta el signo de cada operación.' : ''),
        'Alumnado con NEAE: se puede ampliar el tiempo, leer los ítems en voz alta y usar material manipulativo en la Parte 2 (anótalo en observaciones).']),
      H('2 · Corrección', COL, 12),
      P('**Parte 1:** 10 ítems × 1 punto. La nota es el número de aciertos.', { sz: 11 }),
      P('**Parte 2:** 6 operaciones. Correcta = 1 punto; proceso bien hecho con un solo error de cálculo = 0,5; mal planteada o en blanco = 0.' + (hayDiv ? ' En las divisiones, cociente y resto correctos = 1; solo el cociente correcto = 0,5.' : ''), { sz: 11 }),
      P('Conversión de la Parte 2 a nota:', { sz: 10.5, c: '#555555', after: 60 }),
      pwNotaTabla(6, 0.5),
      H('3 · Cómo se meten los datos en la app', COL, 12),
      ...lista([
        `Entra en **Registro**, curso **${cursoId}**, **${m.corto}**, y pulsa sobre las siglas del alumno/a.`,
        'En **Cálculo** pulsa **＋ Añadir nota** dos veces: **Nota 1 = Parte 1** (cálculo mental) y **Nota 2 = Parte 2** (operaciones).',
        'La app calcula la media del trimestre. Si haces más pruebas de cálculo en el trimestre, añádelas como Nota 3, Nota 4…']),
    ]));
    const solM = [['Nº', 'Ítem', 'Solución']].concat(pr.mental.map((it, i) => [String(i + 1), it.q, `**${it.r}**`]));
    const solE = [['', 'Operación', 'Solución']].concat(pr.escrito.map((o, i) => {
      const enun = o.t === 'v' ? `${o.a} ${o.s} ${o.b}` : o.t === 'd' ? `${o.a} : ${o.b}` : o.q.replace(/=\s*$/, '');
      return [String.fromCharCode(97 + i) + ')', enun, `**${o.r}**`];
    }));
    const notas = [];
    if(hayDiv) notas.push('c = cociente · r = resto.');
    if(pr.escrito.some(o => /\//.test(o.r)) || pr.mental.some(o => /\//.test(o.r))) notas.push('En las fracciones vale cualquier fracción equivalente; simplificar no es obligatorio.');
    hojas.push(hoja(cab(m), [
      H(`Solucionario · ${base} · ${m.corto}`, COL, 13, { before: 0 }),
      H('Parte 1 · Cálculo mental', COL, 12),
      T(solM, { anchos: [700, 6164, 3000], cabecera: '#FFF4E0', celda: { sz: 11 } }),
      H('Parte 2 · Operaciones', COL, 12, { before: 200 }),
      T(solE, { anchos: [700, 6164, 3000], cabecera: '#FFF4E0', celda: { sz: 11 } }),
      ...notas.map(n => P(n, { sz: 9.5, c: '#777777', before: 80 })),
    ]));
    hojas.push(pwRegistro(cursoId, m, COL, 'Plan de Matemáticas · Prueba de cálculo',
      ['Parte 1 (aciertos)', 'Parte 2 (puntos)', 'Nota 1', 'Nota 2', 'Observaciones']));
  });
  const quien = tipo === 'alumno' ? 'alumnado' : 'maestro';
  await pwDescargar(`Calculo_${cursoId.replace(/[^\w]/g, '')}_${pwNombreTrim(keys)}_${quien}_${S.anio}.docx`,
    `Pruebas de cálculo ${cursoId} · ${quien}`, hojas);
}

// ══════════════════════════════════════════════════════════════
// RESOLUCIÓN DE PROBLEMAS
// ══════════════════════════════════════════════════════════════
async function generarProblemasWord(cursoId, keys, tipo){
  await pwPreparar();
  const base = cursoBase(cursoId);
  if(!base || !PRUEBAS_RP[base]){ alert('No hay prueba de problemas para este curso.'); return; }
  const peque = (base === '1º' || base === '2º');
  const COL = '#dc2626';
  const cab = m => ['Plan de Matemáticas · Resolución de problemas', `${cursoId} · ${m.corto} · ${anioLabel(S.anio)}`, COL];
  const hojas = [];
  pwMomentos(keys).forEach(m => {
    const pr = PRUEBAS_RP[base][m.key];
    if(tipo === 'alumno'){
      const sz = peque ? 15 : 12.5;
      const hueco = peque ? 2100 : 2350;
      const fino = { top: LB('DDDDDD'), bottom: LB('DDDDDD'), left: LB('DDDDDD'), right: LB('DDDDDD') };
      const prob = (p, i) => [
        T([
          [C(P(`**${i + 1}.**  ${p.e}`, { sz, after: 0, line: peque ? 340 : 300 }), { span: 2, w: PW_W, pad: 120 })],
          [C(P(base === '1º' ? 'Dibuja' : 'Datos', { sz: 10, c: '#777777' }), { w: 4932, borders: fino }), C(P('Operación', { sz: 10, c: '#777777' }), { w: 4932, borders: fino })],
          [C(P('Solución: ____________________________________________________', { sz: peque ? 13 : 11, after: 0 }), { span: 2, w: PW_W, pad: 120 })],
        ], { anchos: [4932, 4932], color: 'BBBBBB', altos: [0, hueco, 0] }),
        P('', { after: 140 }),
      ];
      [['p1', 'Parte 1'], ['p2', 'Parte 2']].forEach(([k, lbl]) => {
        const bloques = [];
        pr[k].forEach((p, i) => {
          const [tabla, sep] = prob(p, i);
          bloques.push(tabla, sep);
        });
        hojas.push(hoja(cab(m), [...pwTitulo(`Problemas · ${lbl}`, 'Resolución de problemas'), ...pwDatosAlumno(),
          P(base === '1º' ? 'Escucha el problema. Dibuja, haz la operación y escribe la solución.' : 'Lee despacio cada problema. Escribe los datos, la operación y la solución con una frase.', { sz: peque ? 13 : 11, c: '#444444', after: 140 }),
          ...bloques]));
      });
      return;
    }
    // ── maestro/a ──
    hojas.push(hoja(cab(m), [
      H(`Guía del maestro/a · Resolución de problemas · ${base} · ${m.corto}`, COL, 13, { before: 0 }),
      P(`**Qué mide:** ${pr.contenidos} La Parte 1 tiene problemas de un paso y la Parte 2, de dos pasos o con más razonamiento.`, { sz: 10.5, c: '#555555' }),
      H('1 · Cómo se aplica', COL, 12),
      ...lista([
        `Colectiva, en el aula, en **dos sesiones** (un día cada parte). Unos ${peque ? '25' : '30'} minutos por parte.`,
        base === '1º' ? '**Lee tú en voz alta cada problema dos veces** antes de que lo resuelvan. Pueden dibujar y contar.'
          : base === '2º' ? 'En el 1er trimestre, **lee tú los problemas en voz alta**. Después, que los lean ellos; puedes releerlos a quien lo pida.'
          : 'El alumnado lee solo. Se puede leer en voz alta a quien tenga dificultades lectoras (anótalo en observaciones): aquí medimos el razonamiento, no la lectura.',
        'No se ayuda a elegir la operación. Sí se puede aclarar una palabra que no entiendan.']),
      H('2 · Corrección (cada parte por separado)', COL, 12),
      P('Cada problema vale **4 puntos**, uno por cada paso. Cada parte suma 12 puntos.', { sz: 11 }),
      T([['Paso', '1 punto si…'],
         ['**Comprende**', base === '1º' ? 'El dibujo representa bien la situación.' : 'Escribe los datos útiles y sabe qué le preguntan.'],
         ['**Planifica**', 'Elige la operación u operaciones correctas.'],
         ['**Calcula**', 'Hace bien la operación (un despiste pequeño: 0,5).'],
         ['**Responde**', peque ? 'Da el resultado con su unidad (lápices, euros…). En 1º vale el número junto al dibujo.' : 'Escribe la solución con una frase completa y la unidad.']],
        { anchos: [2000, 7864], cabecera: '#FDECEC', celda: { sz: 10.5 } }),
      P('Si el resultado es correcto por otro camino válido, se da la puntuación completa.', { sz: 10, c: '#555555', before: 80 }),
      P('Conversión de cada parte a nota:', { sz: 10.5, c: '#555555', after: 60 }),
      pwNotaTabla(12, 1),
      H('3 · Cómo se meten los datos en la app', COL, 12),
      ...lista([
        `Entra en **Registro**, curso **${cursoId}**, **${m.corto}**, y pulsa sobre las siglas del alumno/a.`,
        'En **Resolución de Problemas** pulsa **＋ Añadir nota** dos veces: **Nota 1 = Parte 1** y **Nota 2 = Parte 2**.',
        'La app calcula la media del trimestre.']),
    ]));
    const filas = [['Nº', 'Problema', 'Operaciones', 'Solución']];
    const sep = t => [C(P(t, { sz: 10, b: true, after: 0 }), { span: 4, bg: '#F3F4F6', w: PW_W })];
    filas.push(sep('Parte 1 · problemas de un paso'));
    pr.p1.forEach((p, i) => filas.push([String(i + 1), C(P(p.e, { sz: 9.5, after: 0 }), { w: 4364 }), p.op, `**${p.s}**`]));
    filas.push(sep('Parte 2 · problemas de dos pasos'));
    pr.p2.forEach((p, i) => filas.push([String(i + 1), C(P(p.e, { sz: 9.5, after: 0 }), { w: 4364 }), p.op, `**${p.s}**`]));
    hojas.push(hoja(cab(m), [
      H(`Solucionario · ${base} · ${m.corto}`, COL, 13, { before: 0 }),
      T(filas, { anchos: [500, 4364, 2500, 2500], cabecera: '#FDECEC', celda: { sz: 10 } }),
    ]));
    hojas.push(pwRegistro(cursoId, m, COL, 'Plan de Matemáticas · Resolución de problemas',
      ['Parte 1 (puntos /12)', 'Parte 2 (puntos /12)', 'Nota 1', 'Nota 2', 'Observaciones']));
  });
  const quien = tipo === 'alumno' ? 'alumnado' : 'maestro';
  await pwDescargar(`Problemas_${cursoId.replace(/[^\w]/g, '')}_${pwNombreTrim(keys)}_${quien}_${S.anio}.docx`,
    `Pruebas de problemas ${cursoId} · ${quien}`, hojas);
}
