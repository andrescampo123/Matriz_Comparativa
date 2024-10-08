#ruta archivo
file.choose()

install.packages("readxl")
install.packages("openxlsx")
library(readxl)
library(openxlsx)
#ELIMINAR todso lo que esta entre parentesis 
OBSERVADO2023R <- read_excel("C:\\Users\\ASUS\\Desktop\\Andres y Laura\\INS\\Productos a entregar\\Matriz comparativa\\R\\Resultados R\\DENGUE2024\\OBSERVADO2024R.xlsx", sheet = "Hoja1")

# Eliminar el contenido dentro de paréntesis en la columna MUNICIPIOS
OBSERVADO2023R$MUNICIPIOS <- gsub("\\s*\\([^\\)]+\\)", "", OBSERVADO2023R$MUNICIPIOS)
# sobrescribir el original
write.xlsx(OBSERVADO2023R, "C:\\Users\\ASUS\\Desktop\\Andres y Laura\\INS\\Productos a entregar\\Matriz comparativa\\R\\Resultados R\\DENGUE2024\\OBSERVADO2024R.xlsx")



# Leer la base de datos desde el archivo Excel existente
ruta_archivo <- "C:\\Users\\ASUS\\Desktop\\Andres y Laura\\INS\\Productos a entregar\\Matriz comparativa\\R\\Resultados R\\DENGUE2024\\OBSERVADO2024R.xlsx"


datos <- read_excel(ruta_archivo, sheet = "Sheet 1")

# Crear la base de datos de los códigos de departamentos y municipios
# Aquí se asume que tienes los datos de departamentos y municipios con sus códigos.
departamentos <- c("ANTIOQUIA", "ATLÁNTICO", "BOGOTÁ D.C.", "BOLÍVAR", "BOYACÁ", 
                   "CALDAS", "CAQUETÁ", "CAUCA", "CESAR", "CÓRDOBA", 
                   "CUNDINAMARCA", "CHOCÓ", "HUILA", "LA GUAJIRA", "MAGDALENA", 
                   "META", "NARIÑO", "NORTE DE SANTANDER", "QUINDÍO", "RISARALDA", 
                   "SANTANDER", "SUCRE", "TOLIMA", "VALLE DEL CAUCA", "ARAUCA", 
                   "CASANARE", "PUTUMAYO", "SAN ANDRÉS", "AMAZONAS", "GUAINÍA", 
                   "GUAVIARE", "VAUPÉS", "VICHADA")

codigos_departamentos <- c("05", "08", "11", "13", "15", 
                           "17", "18", "19", "20", "23", 
                           "25", "27", "41", "44", "47", 
                           "50", "52", "54", "63", "66", 
                           "68", "70", "73", "76", "81", 
                           "85", "86", "88", "91", "94", 
                           "95", "97", "99")

# Crear una tabla de departamentos y sus códigos
tabla_departamentos <- data.frame(DEPARTAMENTOS = departamentos, 
                                  CODIGO_DEPARTAMENTO = codigos_departamentos)

# Lista de códigos de municipios
codigos_municipios <- c(
  "05001", "05002", "05004", "05021", "05030", "05031", "05034", "05036", "05038", "05040", "05042", 
  "05044", "05045", "05051", "05055", "05059", "05079", "05086", "05088", "05091", "05093", "05101", 
  "05107", "05113", "05120", "05125", "05129", "05134", "05138", "05142", "05145", "05147", "05148", 
  "05150", "05154", "05172", "05190", "05197", "05206", "05209", "05212", "05234", "05237", "05240", 
  "05250", "05264", "05266", "05282", "05284", "05306", "05308", "05310", "05313", "05315", "05318", 
  "05321", "05347", "05353", "05360", "05361", "05364", "05368", "05376", "05380", "05390", "05400", 
  "05411", "05425", "05440", "05467", "05475", "05480", "05483", "05490", "05495", "05501", "05541", 
  "05543", "05576", "05579", "05585", "05591", "05604", "05607", "05615", "05628", "05631", "05642", 
  "05647", "05649", "05652", "05656", "05658", "05659", "05660", "05664", "05665", "05667", "05670", 
  "05674", "05679", "05686", "05690", "05697", "05736", "05756", "05761", "05789", "05790", "05792", 
  "05809", "05819", "05837", "05842", "05847", "05854", "05856", "05858", "05861", "05873", "05885", 
  "05887", "05890", "05893", "05895", "08001", "08078", "08137", "08141", "08296", "08372", "08421", 
  "08433", "08436", "08520", "08549", "08558", "08560", "08573", "08606", "08634", "08638", "08675", 
  "08685", "08758", "08770", "08832", "08849", "11001", "13001", "13006", "13030", "13042", "13052", 
  "13062", "13074", "13140", "13160", "13188", "13212", "13222", "13244", "13248", "13268", "13300", 
  "13430", "13433", "13440", "13442", "13458", "13468", "13473", "13549", "13580", "13600", "13620", 
  "13647", "13650", "13654", "13655", "13657", "13667", "13670", "13673", "13683", "13688", "13744", 
  "13760", "13780", "13810", "13836", "13838", "13873", "13894", "15001", "15022", "15047", "15051", 
  "15087", "15090", "15092", "15097", "15104", "15106", "15109", "15114", "15131", "15135", "15162", 
  "15172", "15176", "15180", "15183", "15185", "15187", "15189", "15204", "15212", "15215", "15218", 
  "15223", "15224", "15226", "15232", "15236", "15238", "15244", "15248", "15272", "15276", "15293", 
  "15296", "15299", "15317", "15322", "15325", "15332", "15362", "15367", "15368", "15377", "15380", 
  "15401", "15403", "15407", "15425", "15442", "15455", "15464", "15466", "15469", "15476", "15480", 
  "15491", "15494", "15500", "15507", "15511", "15514", "15516", "15518", "15522", "15531", "15533", 
  "15537", "15542", "15550", "15572", "15580", "15599", "15600", "15621", "15632", "15638", "15646", 
  "15660", "15664", "15667", "15673", "15676", "15681", "15686", "15690", "15693", "15696", "15720", 
  "15723", "15740", "15753", "15755", "15757", "15759", "15761", "15762", "15763", "15764", "15774", 
  "15776", "15778", "15790", "15798", "15804", "15806", "15808", "15810", "15814", "15816", "15820", 
  "15822", "15832", "15835", "15837", "15839", "15842", "15861", "15879", "15897", "17001", "17013", 
  "17042", "17050", "17088", "17174", "17272", "17380", "17388", "17433", "17442", "17444", "17446", 
  "17486", "17495", "17513", "17524", "17541", "17614", "17616", "17653", "17662", "17665", "17777", 
  "17867", "17873", "17877", "18001", "18029", "18094", "18150", "18205", "18247", "18256", "18410", 
  "18460", "18479", "18592", "18610", "18753", "18756", "18785", "18860", "19001", "19022", "19050", 
  "19075", "19100", "19110", "19130", "19137", "19142", "19212", "19256", "19290", "19300", "19318", 
  "19355", "19364", "19392", "19397", "19418", "19450", "19455", "19473", "19513", "19517", "19532", 
  "19533", "19548", "19573", "19585", "19622", "19693", "19698", "19701", "19743", "19760", "19780", 
  "19785", "19807", "19809", "19821", "19824", "19845", "20001", "20011", "20013", "20032", "20045", 
  "20060", "20175", "20178", "20228", "20238", "20250", "20295", "20310", "20383", "20400", "20443", 
  "20517", "20550", "20570", "20614", "20621", "20710", "20750", "20770", "20787", "23001", "23068", 
  "23079", "23090", "23162", "23168", "23182", "23189", "23300", "23350", "23417", "23419", "23464", 
  "23466", "23500", "23555", "23570", "23574", "23580", "23586", "23660", "23670", "23672", "23675", 
  "23678", "23686", "23807", "23855", "25001", "25019", "25035", "25040", "25053", "25086", "25095", 
  "25099", "25120", "25123", "25126", "25148", "25151", "25154", "25168", "25175", "25178", "25181", 
  "25183", "25200", "25214", "25224", "25245", "25258", "25260", "25269", "25279", "25281", "25286", 
  "25288", "25290", "25293", "25295", "25297", "25299", "25307", "25312", "25317", "25320", "25322", 
  "25324", "25326", "25328", "25335", "25339", "25368", "25372", "25377", "25386", "25394", "25398", 
  "25402", "25407", "25426", "25430", "25436", "25438", "25473", "25483", "25486", "25488", "25489", 
  "25491", "25506", "25513", "25518", "25524", "25530", "25535", "25572", "25580", "25592", "25594", 
  "25596", "25599", "25612", "25645", "25649", "25653", "25658", "25662", "25718", "25736", "25740", 
  "25743", "25745", "25754", "25758", "25769", "25772", "25777", "25779", "25781", "25785", "25793", 
  "25797", "25799", "25805", "25807", "25815", "25817", "25823", "25839", "25841", "25843", "25845", 
  "25851", "25862", "25867", "25871", "25873", "25875", "25878", "25885", "25898", "25899", "27001", 
  "27006", "27025", "27050", "27073", "27075", "27077", "27086", "27099", "27135", "27150", "27160", 
  "27205", "27245", "27250", "27361", "27372", "27413", "27425", "27430", "27450", "27491", "27495", 
  "27580", "27600", "27615", "27660", "27745", "27787", "27800", "27810", "41001", "41006", "41013", 
  "41016", "41020", "41026", "41078", "41132", "41206", "41244", "41298", "41306", "41319", "41349", 
  "41357", "41359", "41378", "41396", "41483", "41503", "41518", "41524", "41530", "41548", "41551", 
  "41615", "41660", "41668", "41676", "41770", "41791", "41797", "41799", "41801", "41807", "41872", 
  "41885", "44001", "44035", "44078", "44090", "44098", "44110", "44279", "44378", "44420", "44430", 
  "44560", "44650", "44847", "44855", "44874", "47001", "47030", "47053", "47058", "47161", "47170", 
  "47189", "47205", "47245", "47258", "47268", "47288", "47318", "47460", "47541", "47545", "47551", 
  "47555", "47570", "47605", "47660", "47675", "47692", "47703", "47707", "47720", "47745", "47798", 
  "47960", "47980", "50001", "50006", "50110", "50124", "50150", "50223", "50226", "50245", "50251", 
  "50270", "50287", "50313", "50318", "50325", "50330", "50350", "50370", "50400", "50450", "50568", 
  "50573", "50577", "50590", "50606", "50680", "50683", "50686", "50689", "50711", "52001", "52019", 
  "52022", "52036", "52051", "52079", "52083", "52110", "52203", "52207", "52210", "52215", "52224", 
  "52227", "52233", "52240", "52250", "52254", "52256", "52258", "52260", "52287", "52317", "52320", 
  "52323", "52352", "52354", "52356", "52378", "52381", "52385", "52390", "52399", "52405", "52411", 
  "52418", "52427", "52435", "52473", "52480", "52490", "52506", "52520", "52540", "52560", "52565", 
  "52573", "52585", "52612", "52621", "52678", "52683", "52685", "52687", "52693", "52694", "52696", 
  "52699", "52720", "52786", "52788", "52835", "52838", "52885", "54001", "54003", "54051", "54099", 
  "54109", "54125", "54128", "54172", "54174", "54206", "54223", "54239", "54245", "54250", "54261", 
  "54313", "54344", "54347", "54377", "54385", "54398", "54405", "54418", "54480", "54498", "54518", 
  "54520", "54553", "54599", "54660", "54670", "54673", "54680", "54720", "54743", "54800", "54810", 
  "54820", "54871", "54874", "63001", "63111", "63130", "63190", "63212", "63272", "63302", "63401", 
  "63470", "63548", "63594", "63690", "66001", "66045", "66075", "66088", "66170", "66318", "66383", 
  "66400", "66440", "66456", "66572", "66594", "66682", "66687", "68001", "68013", "68020", "68051", 
  "68077", "68079", "68081", "68092", "68101", "68121", "68132", "68147", "68152", "68160", "68162", 
  "68167", "68169", "68176", "68179", "68190", "68207", "68209", "68211", "68217", "68229", "68235", 
  "68245", "68250", "68255", "68264", "68266", "68271", "68276", "68296", "68298", "68307", "68318", 
  "68320", "68322", "68324", "68327", "68344", "68368", "68370", "68377", "68385", "68397", "68406", 
  "68418", "68425", "68432", "68444", "68464", "68468", "68498", "68500", "68502", "68522", "68524", 
  "68533", "68547", "68549", "68572", "68573", "68575", "68615", "68655", "68669", "68673", "68679", 
  "68682", "68684", "68686", "68689", "68705", "68720", "68745", "68755", "68770", "68773", "68780", 
  "68820", "68855", "68861", "68867", "68872", "68895", "70001", "70110", "70124", "70204", "70215", 
  "70221", "70230", "70233", "70235", "70265", "70400", "70418", "70429", "70473", "70508", "70523", 
  "70670", "70678", "70702", "70708", "70713", "70717", "70742", "70771", "70820", "70823", "73001", 
  "73024", "73026", "73030", "73043", "73055", "73067", "73124", "73148", "73152", "73168", "73200", 
  "73217", "73226", "73236", "73268", "73270", "73275", "73283", "73319", "73347", "73349", "73352", 
  "73408", "73411", "73443", "73449", "73461", "73483", "73504", "73520", "73547", "73555", "73563", 
  "73585", "73616", "73622", "73624", "73671", "73675", "73678", "73686", "73770", "73854", "73861", 
  "73870", "73873", "76001", "76020", "76036", "76041", "76054", "76100", "76109", "76111", "76113", 
  "76122", "76126", "76130", "76147", "76233", "76243", "76246", "76248", "76250", "76275", "76306", 
  "76318", "76364", "76377", "76400", "76403", "76497", "76520", "76563", "76606", "76616", "76622", 
  "76670", "76736", "76823", "76828", "76834", "76845", "76863", "76869", "76890", "76892", "76895", 
  "81001", "81065", "81220", "81300", "81591", "81736", "81794", "85001", "85010", "85015", "85125", 
  "85136", "85139", "85162", "85225", "85230", "85250", "85263", "85279", "85300", "85315", "85325", 
  "85400", "85410", "85430", "85440", "86001", "86219", "86320", "86568", "86569", "86571", "86573", 
  "86749", "86755", "86757", "86760", "86865", "86885", "88001", "88564", "91001", "91263", "91405", 
  "91407", "91430", "91460", "91530", "91536", "91540", "91669", "91798", "94001", "94343", "94663", 
  "94883", "94884", "94885", "94886", "94887", "94888", "95001", "95015", "95025", "95200", "97001", 
  "97161", "97511", "97666", "97777", "97889", "99001", "99524", "99624", "99773","05321","13490","23682","23815"
)

# Lista de municipios correspondientes
municipios <- c(
  "MEDELLÍN", "ABEJORRAL", "ABRIAQUÍ", "ALEJANDRÍA", "AMAGÁ", "AMALFI", "ANDES", "ANGELÓPOLIS", "ANGOSTURA", "ANORÍ", 
  "SANTAFÉ DE ANTIOQUIA", "ANZA", "APARTADÓ", "ARBOLETES", "ARGELIA", "ARMENIA", "BARBOSA", "BELMIRA", "BELLO", "BETANIA", 
  "BETULIA", "CIUDAD BOLÍVAR", "BRICEÑO", "BURITICÁ", "CÁCERES", "CAICEDO", "CALDAS", "CAMPAMENTO", "CAÑASGORDAS", 
  "CARACOLÍ", "CARAMANTA", "CAREPA", "EL CARMEN DE VIBORAL", "CAROLINA", "CAUCASIA", "CHIGORODÓ", "CISNEROS", 
  "COCORNÁ", "CONCEPCIÓN", "CONCORDIA", "COPACABANA", "DABEIBA", "DON MATÍAS", "EBÉJICO", "EL BAGRE", "ENTRERRIOS", 
  "ENVIGADO", "FREDONIA", "FRONTINO", "GIRALDO", "GIRARDOTA", "GÓMEZ PLATA", "GRANADA", "GUADALUPE", "GUARNE", 
  "GUATAPE", "HELICONIA", "HISPANIA", "ITAGUI", "ITUANGO", "JARDÍN", "JERICÓ", "LA CEJA", "LA ESTRELLA", 
  "LA PINTADA", "LA UNIÓN", "LIBORINA", "MACEO", "MARINILLA", "MONTEBELLO", "MURINDÓ", "MUTATÁ", "NARIÑO", 
  "NECOCLÍ", "NECHÍ", "OLAYA", "PEÑOL", "PEQUE", "PUEBLORRICO", "PUERTO BERRÍO", "PUERTO NARE", "PUERTO TRIUNFO", 
  "REMEDIOS", "RETIRO", "RIONEGRO", "SABANALARGA", "SABANETA", "SALGAR", "SAN ANDRÉS DE CUERQUÍA", "SAN CARLOS", 
  "SAN FRANCISCO", "SAN JERÓNIMO", "SAN JOSÉ DE LA MONTAÑA", "SAN JUAN DE URABÁ", "SAN LUIS", "SAN PEDRO", 
  "SAN PEDRO DE URABA", "SAN RAFAEL", "SAN ROQUE", "SAN VICENTE", "SANTA BÁRBARA", "SANTA ROSA DE OSOS", 
  "SANTO DOMINGO", "EL SANTUARIO", "SEGOVIA", "SONSON", "SOPETRÁN", "TÁMESIS", "TARAZÁ", "TARSO", "TITIRIBÍ", 
  "TOLEDO", "TURBO", "URAMITA", "URRAO", "VALDIVIA", "VALPARAÍSO", "VEGACHÍ", "VENECIA", "VIGÍA DEL FUERTE", 
  "YALÍ", "YARUMAL", "YOLOMBÓ", "YONDÓ", "ZARAGOZA", "BARRANQUILLA", "BARANOA", "CAMPO DE LA CRUZ", "CANDELARIA", 
  "GALAPA", "JUAN DE ACOSTA", "LURUACO", "MALAMBO", "MANATÍ", "PALMAR DE VARELA", "PIOJÓ", "POLONUEVO", "PONEDERA", 
  "PUERTO COLOMBIA", "REPELÓN", "SABANAGRANDE", "SABANALARGA", "SANTA LUCÍA", "SANTO TOMÁS", "SOLEDAD", "SUAN", 
  "TUBARÁ", "USIACURÍ", "BOGOTÁ, D.C.", "CARTAGENA", "ACHÍ", "ALTOS DEL ROSARIO", "ARENAL", "ARJONA", 
  "ARROYOHONDO", "BARRANCO DE LOBA", "CALAMAR", "CANTAGALLO", "CICUCO", "CÓRDOBA", "CLEMENCIA", "EL CARMEN DE BOLÍVAR", 
  "EL GUAMO", "EL PEÑÓN", "HATILLO DE LOBA", "MAGANGUÉ", "MAHATES", "MARGARITA", "MARÍA LA BAJA", "MONTECRISTO", 
  "MOMPÓS", "MORALES", "PINILLOS", "REGIDOR", "RÍO VIEJO", "SAN CRISTÓBAL", "SAN ESTANISLAO", "SAN FERNANDO", 
  "SAN JACINTO", "SAN JACINTO DEL CAUCA", "SAN JUAN NEPOMUCENO", "SAN MARTÍN DE LOBA", "SAN PABLO", "SANTA CATALINA", 
  "SANTA ROSA", "SANTA ROSA DEL SUR", "SIMITÍ", "SOPLAVIENTO", "TALAIGUA NUEVO", "TIQUISIO", "TURBACO", 
  "TURBANÁ", "VILLANUEVA", "ZAMBRANO", "TUNJA", "ALMEIDA", "AQUITANIA", "ARCABUCO", "BELÉN", "BERBEO", 
  "BETÉITIVA", "BOAVITA", "BOYACÁ", "BRICEÑO", "BUENAVISTA", "BUSBANZÁ", "CALDAS", "CAMPOHERMOSO", "CERINZA", 
  "CHINAVITA", "CHIQUINQUIRÁ", "CHISCAS", "CHITA", "CHITARAQUE", "CHIVATÁ", "CIÉNEGA", "CÓMBITA", "COPER", 
  "CORRALES", "COVARACHÍA", "CUBARÁ", "CUCAITA", "CUÍTIVA", "CHÍQUIZA", "CHIVOR", "DUITAMA", "EL COCUY", 
  "EL ESPINO", "FIRAVITOBA", "FLORESTA", "GACHANTIVÁ", "GAMEZA", "GARAGOA", "GUACAMAYAS", "GUATEQUE", 
  "GUAYATÁ", "GÜICÁN", "IZA", "JENESANO", "JERICÓ", "LABRANZAGRANDE", "LA CAPILLA", "LA VICTORIA", 
  "LA UVITA", "VILLA DE LEYVA", "MACANAL", "MARIPÍ", "MIRAFLORES", "MONGUA", "MONGUÍ", "MONIQUIRÁ", 
  "MOTAVITA", "MUZO", "NOBSA", "NUEVO COLÓN", "OICATÁ", "OTANCHE", "PACHAVITA", "PÁEZ", "PAIPA", 
  "PAJARITO", "PANQUEBA", "PAUNA", "PAYA", "PAZ DE RÍO", "PESCA", "PISBA", "PUERTO BOYACÁ", "QUÍPAMA", 
  "RAMIRIQUÍ", "RÁQUIRA", "RONDÓN", "SABOYÁ", "SÁCHICA", "SAMACÁ", "SAN EDUARDO", "SAN JOSÉ DE PARE", 
  "SAN LUIS DE GACENO", "SAN MATEO", "SAN MIGUEL DE SEMA", "SAN PABLO DE BORBUR", "SANTANA", "SANTA MARÍA", 
  "SANTA ROSA DE VITERBO", "SANTA SOFÍA", "SATIVANORTE", "SATIVASUR", "SIACHOQUE", "SOATÁ", "SOCOTÁ", 
  "SOCHA", "SOGAMOSO", "SOMONDOCO", "SORA", "SOTAQUIRÁ", "SORACÁ", "SUSACÓN", "SUTAMARCHÁN", "SUTATENZA", 
  "TASCO", "TENZA", "TIBANÁ", "TIBASOSA", "TINJACÁ", "TIPACOQUE", "TOCA", "TOGÜÍ", "TÓPAGA", 
  "TOTA", "TUNUNGUÁ", "TURMEQUÉ", "TUTA", "TUTAZÁ", "UMBITA", "VENTAQUEMADA", "VIRACACHÁ", "ZETAQUIRA", 
  "MANIZALES", "AGUADAS", "ANSERMA", "ARANZAZU", "BELALCÁZAR", "CHINCHINÁ", "FILADELFIA", "LA DORADA", 
  "LA MERCED", "MANZANARES", "MARMATO", "MARQUETALIA", "MARULANDA", "NEIRA", "NORCASIA", "PÁCORA", 
  "PALESTINA", "PENSILVANIA", "RIOSUCIO", "RISARALDA", "SALAMINA", "SAMANÁ", "SAN JOSÉ", "SUPÍA", 
  "VICTORIA", "VILLAMARÍA", "VITERBO", "FLORENCIA", "ALBANIA", "BELÉN DE LOS ANDAQUIES", "CARTAGENA DEL CHAIRÁ", 
  "CURILLO", "EL DONCELLO", "EL PAUJIL", "LA MONTAÑITA", "MILÁN", "MORELIA", "PUERTO RICO", "SAN JOSÉ DEL FRAGUA", 
  "SAN VICENTE DEL CAGUÁN", "SOLANO", "SOLITA", "VALPARAÍSO", "POPAYÁN", "ALMAGUER", "ARGELIA", "BALBOA", 
  "BOLÍVAR", "BUENOS AIRES", "CAJIBÍO", "CALDONO", "CALOTO", "CORINTO", "EL TAMBO", "FLORENCIA", 
  "GUACHENÉ", "GUAPI", "INZÁ", "JAMBALÓ", "LA SIERRA", "LA VEGA", "LÓPEZ", "MERCADERES", 
  "MIRANDA", "MORALES", "PADILLA", "PAEZ", "PATÍA", "PIAMONTE", "PIENDAMÓ", "PUERTO TEJADA", 
  "PURACÉ", "ROSAS", "SAN SEBASTIÁN", "SANTANDER DE QUILICHAO", "SANTA ROSA", "SILVIA", "SOTARA", 
  "SUÁREZ", "SUCRE", "TIMBÍO", "TIMBIQUÍ", "TORIBIO", "TOTORÓ", "VILLA RICA", "VALLEDUPAR", 
  "AGUACHICA", "AGUSTÍN CODAZZI", "ASTREA", "BECERRIL", "BOSCONIA", "CHIMICHAGUA", "CHIRIGUANÁ", "CURUMANÍ", 
  "EL COPEY", "EL PASO", "GAMARRA", "GONZÁLEZ", "LA GLORIA", "LA JAGUA DE IBIRICO", "MANAURE", 
  "PAILITAS", "PELAYA", "PUEBLO BELLO", "RÍO DE ORO", "LA PAZ", "SAN ALBERTO", "SAN DIEGO", "SAN MARTÍN", 
  "TAMALAMEQUE", "MONTERÍA", "AYAPEL", "BUENAVISTA", "CANALETE", "CERETÉ", "CHIMÁ", "CHINÚ", 
  "CIÉNAGA DE ORO", "COTORRA", "LA APARTADA", "LORICA", "LOS CÓRDOBAS", "MOMIL", "MONTELÍBANO", 
  "MOÑITOS", "PLANETA RICA", "PUEBLO NUEVO", "PUERTO ESCONDIDO", "PUERTO LIBERTADOR", "PURÍSIMA", 
  "SAHAGÚN", "SAN ANDRÉS SOTAVENTO", "SAN ANTERO", "SAN BERNARDO DEL VIENTO", "SAN CARLOS", "SAN PELAYO", 
  "TIERRALTA", "VALENCIA", "AGUA DE DIOS", "ALBÁN", "ANAPOIMA", "ANOLAIMA", "ARBELÁEZ", "BELTRÁN", 
  "BITUIMA", "BOJACÁ", "CABRERA", "CACHIPAY", "CAJICÁ", "CAPARRAPÍ", "CAQUEZA", "CARMEN DE CARUPA", 
  "CHAGUANÍ", "CHÍA", "CHIPAQUE", "CHOACHÍ", "CHOCONTÁ", "COGUA", "COTA", "CUCUNUBÁ", 
  "EL COLEGIO", "EL PEÑÓN", "EL ROSAL", "FACATATIVÁ", "FOMEQUE", "FOSCA", "FUNZA", "FÚQUENE", 
  "FUSAGASUGÁ", "GACHALA", "GACHANCIPÁ", "GACHETÁ", "GAMA", "GIRARDOT", "GRANADA", "GUACHETÁ", 
  "GUADUAS", "GUASCA", "GUATAQUÍ", "GUATAVITA", "GUAYABAL DE SIQUIMA", "GUAYABETAL", "GUTIÉRREZ", 
  "JERUSALÉN", "JUNÍN", "LA CALERA", "LA MESA", "LA PALMA", "LA PEÑA", "LA VEGA", "LENGUAZAQUE", 
  "MACHETA", "MADRID", "MANTA", "MEDINA", "MOSQUERA", "NARIÑO", "NEMOCÓN", "NILO", 
  "NIMAIMA", "NOCAIMA", "VENECIA", "PACHO", "PAIME", "PANDI", "PARATEBUENO", "PASCA", 
  "PUERTO SALGAR", "PULÍ", "QUEBRADANEGRA", "QUETAME", "QUIPILE", "APULO", "RICAURTE", "SAN ANTONIO DEL TEQUENDAMA", 
  "SAN BERNARDO", "SAN CAYETANO", "SAN FRANCISCO", "SAN JUAN DE RÍO SECO", "SASAIMA", "SESQUILÉ", "SIBATÉ", 
  "SILVANIA", "SIMIJACA", "SOACHA", "SOPÓ", "SUBACHOQUE", "SUESCA", "SUPATÁ", "SUSA", 
  "SUTATAUSA", "TABIO", "TAUSA", "TENA", "TENJO", "TIBACUY", "TIBIRITA", "TOCAIMA", 
  "TOCANCIPÁ", "TOPAIPÍ", "UBALÁ", "UBAQUE", "VILLA DE SAN DIEGO DE UBATE", "UNE", "ÚTICA", "VERGARA", 
  "VIANÍ", "VILLAGÓMEZ", "VILLAPINZÓN", "VILLETA", "VIOTÁ", "YACOPÍ", "ZIPACÓN", "ZIPAQUIRÁ", 
  "QUIBDÓ", "ACANDÍ", "ALTO BAUDO", "ATRATO", "BAGADÓ", "BAHÍA SOLANO", "BAJO BAUDÓ", "BELÉN DE BAJIRÁ", 
  "BOJAYA", "EL CANTÓN DEL SAN PABLO", "CARMEN DEL DARIEN", "CÉRTEGUI", "CONDOTO", "EL CARMEN DE ATRATO", "EL LITORAL DEL SAN JUAN", 
  "ISTMINA", "JURADÓ", "LLORÓ", "MEDIO ATRATO", "MEDIO BAUDÓ", "MEDIO SAN JUAN", "NÓVITA", "NUQUÍ", 
  "RÍO IRO", "RÍO QUITO", "RIOSUCIO", "SAN JOSÉ DEL PALMAR", "SIPÍ", "TADÓ", "UNGUÍA", "UNIÓN PANAMERICANA", 
  "NEIVA", "ACEVEDO", "AGRADO", "AIPE", "ALGECIRAS", "ALTAMIRA", "BARAYA", "CAMPOALEGRE", 
  "COLOMBIA", "ELÍAS", "GARZÓN", "GIGANTE", "GUADALUPE", "HOBO", "IQUIRA", "ISNOS", 
  "LA ARGENTINA", "LA PLATA", "NÁTAGA", "OPORAPA", "PAICOL", "PALERMO", "PALESTINA", "PITAL", 
  "PITALITO", "RIVERA", "SALADOBLANCO", "SAN AGUSTÍN", "SANTA MARÍA", "SUAZA", "TARQUI", "TESALIA", 
  "TELLO", "TERUEL", "TIMANÁ", "VILLAVIEJA", "YAGUARÁ", "RIOHACHA", "ALBANIA", "BARRANCAS", 
  "DIBULLA", "DISTRACCIÓN", "EL MOLINO", "FONSECA", "HATONUEVO", "LA JAGUA DEL PILAR", "MAICAO", 
  "MANAURE", "SAN JUAN DEL CESAR", "URIBIA", "URUMITA", "VILLANUEVA", "SANTA MARTA", "ALGARROBO", 
  "ARACATACA", "ARIGUANÍ", "CERRO SAN ANTONIO", "CHIBOLO", "CIÉNAGA", "CONCORDIA", "EL BANCO", 
  "EL PIÑON", "EL RETÉN", "FUNDACIÓN", "GUAMAL", "NUEVA GRANADA", "PEDRAZA", "PIJIÑO DEL CARMEN", 
  "PIVIJAY", "PLATO", "PUEBLOVIEJO", "REMOLINO", "SABANAS DE SAN ANGEL", "SALAMINA", "SAN SEBASTIÁN DE BUENAVISTA", 
  "SAN ZENÓN", "SANTA ANA", "SANTA BÁRBARA DE PINTO", "SITIONUEVO", "TENERIFE", "ZAPAYÁN", "ZONA BANANERA", 
  "VILLAVICENCIO", "ACACÍAS", "BARRANCA DE UPÍA", "CABUYARO", "CASTILLA LA NUEVA", "CUBARRAL", "CUMARAL", 
  "EL CALVARIO", "EL CASTILLO", "EL DORADO", "FUENTE DE ORO", "GRANADA", "GUAMAL", "MAPIRIPÁN", 
  "MESETAS", "LA MACARENA", "URIBE", "LEJANÍAS", "PUERTO CONCORDIA", "PUERTO GAITÁN", "PUERTO LÓPEZ", 
  "PUERTO LLERAS", "PUERTO RICO", "RESTREPO", "SAN CARLOS DE GUAROA", "SAN JUAN DE ARAMA", "SAN JUANITO", 
  "SAN MARTÍN", "VISTAHERMOSA", "PASTO", "ALBÁN", "ALDANA", "ANCUYÁ", "ARBOLEDA", "BARBACOAS", 
  "BELÉN", "BUESACO", "COLÓN", "CONSACA", "CONTADERO", "CÓRDOBA", "CUASPUD", "CUMBAL", 
  "CUMBITARA", "CHACHAGÜÍ", "EL CHARCO", "EL PEÑOL", "EL ROSARIO", "EL TABLÓN DE GÓMEZ", "EL TAMBO", 
  "FUNES", "GUACHUCAL", "GUAITARILLA", "GUALMATÁN", "ILES", "IMUÉS", "IPIALES", "LA CRUZ", 
  "LA FLORIDA", "LA LLANADA", "LA TOLA", "LA UNIÓN", "LEIVA", "LINARES", "LOS ANDES", 
  "MAGÜI", "MALLAMA", "MOSQUERA", "NARIÑO", "OLAYA HERRERA", "OSPINA", "FRANCISCO PIZARRO", 
  "POLICARPA", "POTOSÍ", "PROVIDENCIA", "PUERRES", "PUPIALES", "RICAURTE", "ROBERTO PAYÁN", 
  "SAMANIEGO", "SANDONÁ", "SAN BERNARDO", "SAN LORENZO", "SAN PABLO", "SAN PEDRO DE CARTAGO", "SANTA BÁRBARA", 
  "SANTACRUZ", "SAPUYES", "TAMINANGO", "TANGUA", "SAN ANDRES DE TUMACO", "TÚQUERRES", "YACUANQUER", 
  "CÚCUTA", "ABREGO", "ARBOLEDAS", "BOCHALEMA", "BUCARASICA", "CÁCOTA", "CACHIRÁ", "CHINÁCOTA", 
  "CHITAGÁ", "CONVENCIÓN", "CUCUTILLA", "DURANIA", "EL CARMEN", "EL TARRA", "EL ZULIA", 
  "GRAMALOTE", "HACARÍ", "HERRÁN", "LABATECA", "LA ESPERANZA", "LA PLAYA", "LOS PATIOS", 
  "LOURDES", "MUTISCUA", "OCAÑA", "PAMPLONA", "PAMPLONITA", "PUERTO SANTANDER", "RAGONVALIA", 
  "SALAZAR", "SAN CALIXTO", "SAN CAYETANO", "SANTIAGO", "SARDINATA", "SILOS", "TEORAMA", 
  "TIBÚ", "TOLEDO", "VILLA CARO", "VILLA DEL ROSARIO", "ARMENIA", "BUENAVISTA", "CALARCA", 
  "CIRCASIA", "CÓRDOBA", "FILANDIA", "GÉNOVA", "LA TEBAIDA", "MONTENEGRO", "PIJAO", 
  "QUIMBAYA", "SALENTO", "PEREIRA", "APÍA", "BALBOA", "BELÉN DE UMBRÍA", "DOSQUEBRADAS", 
  "GUÁTICA", "LA CELIA", "LA VIRGINIA", "MARSELLA", "MISTRATÓ", "PUEBLO RICO", "QUINCHÍA", 
  "SANTA ROSA DE CABAL", "SANTUARIO", "BUCARAMANGA", "AGUADA", "ALBANIA", "ARATOCA", 
  "BARBOSA", "BARICHARA", "BARRANCABERMEJA", "BETULIA", "BOLÍVAR", "CABRERA", "CALIFORNIA", 
  "CAPITANEJO", "CARCASÍ", "CEPITÁ", "CERRITO", "CHARALÁ", "CHARTA", "CHIMA", "CHIPATÁ", 
  "CIMITARRA", "CONCEPCIÓN", "CONFINES", "CONTRATACIÓN", "COROMORO", "CURITÍ", "EL CARMEN DE CHUCURÍ", 
  "EL GUACAMAYO", "EL PEÑÓN", "EL PLAYÓN", "ENCINO", "ENCISO", "FLORIÁN", "FLORIDABLANCA", 
  "GALÁN", "GAMBITA", "GIRÓN", "GUACA", "GUADALUPE", "GUAPOTÁ", "GUAVATÁ", "GÜEPSA", 
  "HATO", "JESÚS MARÍA", "JORDÁN", "LA BELLEZA", "LANDÁZURI", "LA PAZ", "LEBRÍJA", 
  "LOS SANTOS", "MACARAVITA", "MÁLAGA", "MATANZA", "MOGOTES", "MOLAGAVITA", "OCAMONTE", 
  "OIBA", "ONZAGA", "PALMAR", "PALMAS DEL SOCORRO", "PÁRAMO", "PIEDECUESTA", "PINCHOTE", 
  "PUENTE NACIONAL", "PUERTO PARRA", "PUERTO WILCHES", "RIONEGRO", "SABANA DE TORRES", "SAN ANDRÉS", 
  "SAN BENITO", "SAN GIL", "SAN JOAQUÍN", "SAN JOSÉ DE MIRANDA", "SAN MIGUEL", "SAN VICENTE DE CHUCURÍ", 
  "SANTA BÁRBARA", "SANTA HELENA DEL OPÓN", "SIMACOTA", "SOCORRO", "SUAITA", "SUCRE", 
  "SURATÁ", "TONA", "VALLE DE SAN JOSÉ", "VÉLEZ", "VETAS", "VILLANUEVA", "ZAPATOCA", 
  "SINCELEJO", "BUENAVISTA", "CAIMITO", "COLOSO", "COROZAL", "COVEÑAS", "CHALÁN", "EL ROBLE", 
  "GALERAS", "GUARANDA", "LA UNIÓN", "LOS PALMITOS", "MAJAGUAL", "MORROA", "OVEJAS", 
  "PALMITO", "SAMPUÉS", "SAN BENITO ABAD", "SAN JUAN DE BETULIA", "SAN MARCOS", "SAN ONOFRE", 
  "SAN PEDRO", "SAN LUIS DE SINCÉ", "SUCRE", "SANTIAGO DE TOLÚ", "TOLÚ VIEJO", "IBAGUÉ", 
  "ALPUJARRA", "ALVARADO", "AMBALEMA", "ANZOÁTEGUI", "ARMERO", "ATACO", "CAJAMARCA", 
  "CARMEN DE APICALÁ", "CASABIANCA", "CHAPARRAL", "COELLO", "COYAIMA", "CUNDAY", "DOLORES", 
  "ESPINAL", "FALAN", "FLANDES", "FRESNO", "GUAMO", "HERVEO", "HONDA", "ICONONZO", 
  "LÉRIDA", "LÍBANO", "MARIQUITA", "MELGAR", "MURILLO", "NATAGAIMA", "ORTEGA", "PALOCABILDO", 
  "PIEDRAS", "PLANADAS", "PRADO", "PURIFICACIÓN", "RIOBLANCO", "RONCESVALLES", "ROVIRA", 
  "SALDAÑA", "SAN ANTONIO", "SAN LUIS", "SANTA ISABEL", "SUÁREZ", "VALLE DE SAN JUAN", 
  "VENADILLO", "VILLAHERMOSA", "VILLARRICA", "CALI", "ALCALÁ", "ANDALUCÍA", "ANSERMANUEVO", 
  "ARGELIA", "BOLÍVAR", "BUENAVENTURA", "GUADALAJARA DE BUGA", "BUGALAGRANDE", "CAICEDONIA", 
  "CALIMA", "CANDELARIA", "CARTAGO", "DAGUA", "EL ÁGUILA", "EL CAIRO", "EL CERRITO", 
  "EL DOVIO", "FLORIDA", "GINEBRA", "GUACARÍ", "JAMUNDÍ", "LA CUMBRE", "LA UNIÓN", 
  "LA VICTORIA", "OBANDO", "PALMIRA", "PRADERA", "RESTREPO", "RIOFRÍO", "ROLDANILLO", 
  "SAN PEDRO", "SEVILLA", "TORO", "TRUJILLO", "TULUÁ", "ULLOA", "VERSALLES", "VIJES", 
  "YOTOCO", "YUMBO", "ZARZAL", "ARAUCA", "ARAUQUITA", "CRAVO NORTE", "FORTUL", "PUERTO RONDÓN", 
  "SARAVENA", "TAME", "YOPAL", "AGUAZUL", "CHAMEZA", "HATO COROZAL", "LA SALINA", "MANÍ", 
  "MONTERREY", "NUNCHÍA", "OROCUÉ", "PAZ DE ARIPORO", "PORE", "RECETOR", "SABANALARGA", 
  "SÁCAMA", "SAN LUIS DE PALENQUE", "TÁMARA", "TAURAMENA", "TRINIDAD", "VILLANUEVA", 
  "MOCOA", "COLÓN", "ORITO", "PUERTO ASÍS", "PUERTO CAICEDO", "PUERTO GUZMÁN", "LEGUÍZAMO", 
  "SIBUNDOY", "SAN FRANCISCO", "SAN MIGUEL", "SANTIAGO", "VALLE DEL GUAMUEZ", "VILLAGARZÓN", 
  "SAN ANDRÉS", "PROVIDENCIA", "LETICIA", "EL ENCANTO", "LA CHORRERA", "LA PEDRERA", 
  "LA VICTORIA", "MIRITI - PARANÁ", "PUERTO ALEGRÍA", "PUERTO ARICA", "PUERTO NARIÑO", 
  "PUERTO SANTANDER", "TARAPACÁ", "INÍRIDA", "BARRANCO MINAS", "MAPIRIPANA", "SAN FELIPE", 
  "PUERTO COLOMBIA", "LA GUADALUPE", "CACAHUAL", "PANA PANA", "MORICHAL", "SAN JOSÉ DEL GUAVIARE", 
  "CALAMAR", "EL RETORNO", "MIRAFLORES", "MITÚ", "CARURU", "PACOA", "TARAIRA", 
  "PAPUNAUA", "YAVARATÉ", "PUERTO CARREÑO", "LA PRIMAVERA", "SANTA ROSALÍA", "CUMARIBO","GUATAPÉ","NOROSÍ","SAN JOSÉ DE URÉ","TUCHÍN"
)



# Crear una tabla de municipios y sus códigos
tabla_municipios <- data.frame(MUNICIPIOS = municipios, 
                               CODIGO_MUNICIPIOS = codigos_municipios)

# Unir los datos originales con los códigos de departamentos
datos <- merge(datos, tabla_departamentos, by = "DEPARTAMENTOS", all.x = TRUE)

# Unir los datos originales con los códigos de municipios
datos <- merge(datos, tabla_municipios, by = "MUNICIPIOS", all.x = TRUE)

# Eliminar duplicados en caso de que existan
datos <- datos[!duplicated(colnames(datos))]

# Reordenar las columnas en el orden deseado
columnas_necesarias <- c("CODIGO_MUNICIPIOS", "MUNICIPIOS", "CODIGO_DEPARTAMENTO", "DEPARTAMENTOS", "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO")
datos <- datos[, columnas_necesarias]
# Escribir el nuevo archivo de Excel con los códigos agregados
write.xlsx(datos, "C:\\Users\\ASUS\\Desktop\\Andres y Laura\\INS\\Productos a entregar\\Matriz comparativa\\R\\Resultados R\\DENGUE2024\\OBSERVADO2024RD.xlsx", sheetName = "Hoja1", overwrite = TRUE)





length(codigos_municipios)
length(municipios)
dif <- which(!(codigos_municipios %in% codigos_municipios[1:length(municipios)]))
codigos_municipios[dif]
setdiff(codigos_municipios, codigos_municipios[1:length(municipios)])
codigos_municipios[length(codigos_municipios)]
codigos_municipios <- codigos_municipios[1:length(municipios)]