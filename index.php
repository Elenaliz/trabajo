<?php
require_once 'phpexcel/Classes/PHPExcel.php';

$objPHPExcel = new PHPExcel();


$objPHPExcel->
	getProperties()
		->setCreator("www.mundoinclusivo.org")
		->setLastModifiedBy("mundoinclusivo.org")
		->setTitle("Datos de Monitoreo")
		->setSubject("Documento exportado")
		->setDescription("Documento generado por Sistema")
		->setKeywords("Usuario Administrador")
		->setCategory("Reportes");

$objPHPExcel->getActiveSheet()->getStyle('A4:EO4')->getFont()->setBold(true)
	->setName("Arial")
	->setSize(8)
	->getColor()->setRGB('000000');
$objPHPExcel->getActiveSheet()->getStyle('A3:EO4')->getFont()->setBold(true)
	->setName("Arial")
	->setSize(8)
	->getColor()->setRGB('000000');



$objPHPExcel->getActiveSheet()->getStyle('M4:X4')->getAlignment()->setTextRotation(90);
$objPHPExcel->getActiveSheet()->getStyle('AC4')->getAlignment()->setTextRotation(90);
$objPHPExcel->getActiveSheet()->getStyle('AH4')->getAlignment()->setTextRotation(90);
$objPHPExcel->getActiveSheet()->getStyle('AU4')->getAlignment()->setTextRotation(90);
$objPHPExcel->getActiveSheet()->getStyle('BI4')->getAlignment()->setTextRotation(90);
$objPHPExcel->getActiveSheet()->getStyle('BJ4:DB4')->getAlignment()->setTextRotation(90);
$objPHPExcel->getActiveSheet()->getStyle('DF4:DN4')->getAlignment()->setTextRotation(90);
$objPHPExcel->getActiveSheet()->getStyle('DP4:EO4')->getAlignment()->setTextRotation(90);

//$objPHPExcel->getActiveSheet()->getStyle('A3:EO3')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_TOP);

$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A4', 'No')
            ->setCellValue('B4', 'CI')
            ->setCellValue('C4', 'Org.')
            ->setCellValue('D4', 'Nombre y Apellido')
            ->setCellValue('E4', 'DIRECCION {Zona, Barrio, Calle N` de casa, canton, comunidad)')
			->setCellValue('F4', 'Fecha de Nac.')
			->setCellValue('G4', 'Edad')
			->setCellValue('H3', 'GENERO')
			->setCellValue('H4', 'M')
			->setCellValue('I4', 'F')
			->setCellValue('J4', 'Fecha de ingreso a la RBC (DD/MM/AA)')
			->setCellValue('K4', 'Situacion  de las acciones de RBC')
			->setCellValue('L4', 'Descripcion medica de la deficiencia')
			->setCellValue('M3', 'TIPO DE DEFICIENCIA Y/O DISCAPACIDAD')
			->setCellValue('M4', 'Deficiencia Visual')
			->setCellValue('N4', 'Deficiencia Auditiva')
			->setCellValue('O4', 'Deficiencia Fisica')
			->setCellValue('P4', 'Sordo Ceguera')
			->setCellValue('Q4', 'Paralisis Cerebral')
			->setCellValue('R4', 'Epilepsia')
			->setCellValue('S4', 'Deficiencia Intelectual')
			->setCellValue('T4', 'Problemas de Aprendizaje')
			->setCellValue('U4', 'Retraso en el Aprendizaje')
			->setCellValue('V4', 'Autismo')
			->setCellValue('W4', 'Deficiencia Psicosocial')
			->setCellValue('X4', 'Multiple')
			->setCellValue('Y3', 'AVANCES TRIMESTRALES')
			->setCellValue('Y4', '1TRIM')
			->setCellValue('Z4', '2TRIM')
			->setCellValue('AA4', '3TRIM')
			->setCellValue('AB4', '4TRIM')
			->setCellValue('AC4', 'VALORACION ANUAL')
			->setCellValue('AD3', 'No DE EVALUACIONES TRIMESTRALES DEL PLAN DE INTERVENCION')
			->setCellValue('AD4', '1TRIM')
			->setCellValue('AE4', '2TRIM')
			->setCellValue('AF4', '3TRIM')
			->setCellValue('AG4', '4TRIM')
			->setCellValue('AH4', 'TOTAL')
			->setCellValue('AI3', 'No DE VISITAS DE MONITOREO TECNICO')
			->setCellValue('AI4', 'E')
			->setCellValue('AJ4', 'F')
			->setCellValue('AK4', 'M')
			->setCellValue('AL4', 'A')
			->setCellValue('AM4', 'M')
			->setCellValue('AN4', 'J')
			->setCellValue('AO4', 'J')
			->setCellValue('AP4', 'A')
			->setCellValue('AQ4', 'S')
			->setCellValue('AR4', 'O')
			->setCellValue('AS4', 'N')
			->setCellValue('AT4', 'D')
			->setCellValue('AU4', 'TOTAL')
			->setCellValue('AV4', 'NOMBRE DE LA PROMOTORA')
			->setCellValue('AW3', 'No DE VISITAS MENSUALES DOMICILIARIAS REALIZADAS POR LA PROMOTORA')
			->setCellValue('AW4', 'E')
			->setCellValue('AX4', 'F')
			->setCellValue('AY4', 'M')
			->setCellValue('AZ4', 'A')
			->setCellValue('BA4', 'M')
			->setCellValue('BB4', 'J')
			->setCellValue('BC4', 'J')
			->setCellValue('BD4', 'A')
			->setCellValue('BE4', 'S')
			->setCellValue('BF4', 'O')
			->setCellValue('BG4', 'N')
			->setCellValue('BH4', 'D')
			->setCellValue('BI4', 'TOTAL')
			->setCellValue('BJ2', 'SALUD')
			->setCellValue('BJ3', 'SERVICIOS MEDICOS')
			->setCellValue('BJ4', 'Visual salud ocular')
			->setCellValue('BK4', 'Tratados vitamina A')
			->setCellValue('BL4', 'Tamizajes visual')
			->setCellValue('BM4', 'Audicion')
			->setCellValue('BN4', 'Tamizajes audicion')
			->setCellValue('BO4', 'Fisicos')
			->setCellValue('BP4', 'Tamizajes fisicos')
			->setCellValue('BQ4', 'Salud mental comunitaria')
			->setCellValue('BR4', 'Otros servicios medicos')
			->setCellValue('BS3', 'CIRUGIAS')
			->setCellValue('BS4', 'Cirugia cataratas')
			->setCellValue('BT4', 'Cirugia Glaucoma')
			->setCellValue('BU4', 'T/Laser retinopatia diabetica')
			->setCellValue('BV4', 'T. Retinopatica del Prematuro')
			->setCellValue('BW4', 'Otras cirugias')
			->setCellValue('BX4', 'C. Mayor O. medio')
			->setCellValue('BY4', 'C. Menor O. Medio')
			->setCellValue('BZ4', 'Otras Cirugias Oido')
			->setCellValue('CA4', 'Otras. C. Otorrinolaringologia')
			->setCellValue('CB4', 'Pies Zambos')
			->setCellValue('CC4', 'Labio leporino P. Hendido')
			->setCellValue('CD4', 'Fistulas Vesicovaginal')
			->setCellValue('CE4', 'C. plastica reconstructiva')
			->setCellValue('CF4', 'Fracturas traumas')
			->setCellValue('CG4', 'Osteiomialitis')
			->setCellValue('CH4', 'Deformidades extremidades')
			->setCellValue('CI4', 'Otras cirigias')
			->setCellValue('CJ3', 'DISPOSITIVOS DE ASISTENCIA')
			->setCellValue('CJ4', 'Lentes/anteojos gafas')
			->setCellValue('CK4', 'Disp/Baja Vision')
			->setCellValue('CL4', 'Audifonos protesis auditivas')
			->setCellValue('CM4', 'Sillas de ruedas triciclos')
			->setCellValue('CN4', 'Ortesis y protesis')
			->setCellValue('CO4', 'Dispositivos para movilidad')
			->setCellValue('CP4', 'Reparacion dispositivos')
			->setCellValue('CQ4', 'Otros dispositovos de asistencia')
			->setCellValue('CR3', 'TERAPIA DE REHABILITACION')
			->setCellValue('CR4', 'Fisioterapia T. ocupacional')
			->setCellValue('CS4', 'Manipulacion pies zambos')
			->setCellValue('CT4', 'Tratamiento no quirurgico otitis media')
			->setCellValue('CU4', 'Audiogramas y pruebas audiologicas')
			->setCellValue('CV4', 'Terapia de lenguaje y Lengua de senas')
			->setCellValue('CW4', 'Desarrollo destresas en baja vision')
			->setCellValue('CX4', 'AVD, incluido autoayuda, orientacion y movilidad ')
			->setCellValue('CY4', 'Terapias/asesoramiento psicosocial')
			->setCellValue('CZ4', 'Otros servicios de rehabilitacion')
			->setCellValue('DA3', 'REFERIR Y FACILITAR OTROS SERVICIOS')
			->setCellValue('DA4', 'Referir servicios medicos y rehabilitacion')
			->setCellValue('DB4', 'Facilitar servicios medicos y rehabilitacion organizados por el proyecto.')
			->setCellValue('DC2', 'EDUCACION')
			->setCellValue('DC4', 'FECHA DE INCLUSION EDUCATIVA')
			->setCellValue('DD4', 'NOMBRE DE LA UNIDAD EDUCATIVA')
			->setCellValue('DE4', 'CURSO')
			->setCellValue('DF3', '1er. Trimestre Escolar')
			->setCellValue('DF4', 'Preparacion para la inclusion')
			->setCellValue('DG4', 'Asesoramiento en Aula ')
			->setCellValue('DH4', 'Grupos de Autoayuda ')
			->setCellValue('DI3', '2do. Trimestre Escolar')
			->setCellValue('DI4', 'Preparacion para la inclusion')
			->setCellValue('DJ4', 'Asesoramiento en Aula ')
			->setCellValue('DK4', 'Grupos de Autoayuda ')
			->setCellValue('DL3', '3er. Trimestre Escolar')
			->setCellValue('DL4', 'Preparacion para la inclusion')
			->setCellValue('DM4', 'Asesoramiento en Aula')
			->setCellValue('DN4', 'Grupos de Autoayuda')
			->setCellValue('DO4', 'NOMBRE DE LOS DOCENTES')
			->setCellValue('DP3', 'EDUCACION')
			->setCellValue('DP4', 'Entornos integrados inclusivos')
			->setCellValue('DQ4', 'Escuela especial clases especiales')
			->setCellValue('DR4', 'Educacion infantil inicial preescolar')
			->setCellValue('DS4', 'Educacion primaria')
			->setCellValue('DT4', 'Educacion segundaria superior')
			->setCellValue('DU4', 'Educacion informal')
			->setCellValue('DV4', 'Cant. PcD graduados educacion segundaria/superior')
			->setCellValue('DW4', 'Cant. PcD recibiendo apoyo en la escula')
			->setCellValue('DX4', 'Cant. PcD recibiendo apoyo en el hogar')
			->setCellValue('DY3', 'SUSTENTO')
			->setCellValue('DY4', 'Cant. Personas recibiendo F. V. largo plazo (mayor 3 meses)')
			->setCellValue('DZ4', 'De las que estan en F. V. convencional')
			->setCellValue('EA4', 'Cant. Personas recibiendo F. V. corto plazo (menor 3 meses)')
			->setCellValue('EB4', 'De las que se han capacitado en situacion incluvivos convencionales')
			->setCellValue('EC4', 'Empleo en el mercado de trabajo formal')
			->setCellValue('ED4', 'Microemprendimiento Auto-empleo')
			->setCellValue('EE4', 'Trabajo en talleres protegidos')
			->setCellValue('EF4', 'Cant. De personas beneficiadas con creditos ofrecidos por instit. Convencionales')
			->setCellValue('EG4', 'De las cuales servicios financieros por instituciones convencionales')
			->setCellValue('EH2', 'INCLUSION / EMPODERAMIENTO')
			->setCellValue('EH3', 'SALVAGUARDA PROTEC. NIÑOS')
			->setCellValue('EH4', 'Personal capacitado en proteccion y salvaguarda de niños')
			->setCellValue('EI4', 'PcD de grupo destinatario participan en taller de proteccion de niñ@s')
			->setCellValue('EJ3', 'INCLUSION DiDRR')
			->setCellValue('EJ4', 'PcD. Capacitadas en reduccion de riegos y desantres')
			->setCellValue('EK3', 'EMPODERAMIENTO (CANT. PARTICIPANTES)')
			->setCellValue('EK4', 'Cant. Participantes en grupos artisticos, culturales deportivos organizados por el proyecto')
			->setCellValue('EL3', 'EMPODERAMIENTO (CANT. DE MIEMBROS)')
			->setCellValue('EL4', 'Cant. De miembros que tienen grupos de niño a niño')
			->setCellValue('EM4', 'Cant. De miembros que tienen grupos de padre a padre')
			->setCellValue('EN4', 'Cant. De miembros que tienen grupos de autoayuda')
			->setCellValue('EO4', 'Cant. De personas que participan en organizaciones de PcD')
			;



$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('0')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('0')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('1')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('1')->setWidth('13');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('2')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('2')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('3')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('3')->setWidth('40');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('4')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('4')->setWidth('20');
//
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('12')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('12')->setWidth('4');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('13')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('13')->setWidth('4');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('14')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('14')->setWidth('4');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('15')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('15')->setWidth('4');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('16')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('16')->setWidth('4');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('17')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('17')->setWidth('4');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('18')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('18')->setWidth('4');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('19')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('19')->setWidth('4');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('20')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('20')->setWidth('4');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('21')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('21')->setWidth('4');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('22')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('22')->setWidth('4');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('23')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('23')->setWidth('4');

$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('24')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('24')->setWidth('6');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('25')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('25')->setWidth('6');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('26')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('26')->setWidth('6');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('27')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('27')->setWidth('6');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('28')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('28')->setWidth('6');

$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('29')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('29')->setWidth('6');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('30')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('30')->setWidth('6');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('31')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('31')->setWidth('6');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('32')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('32')->setWidth('6');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('33')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('33')->setWidth('6');

$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('34')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('34')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('35')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('35')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('36')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('36')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('37')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('37')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('38')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('38')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('39')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('39')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('40')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('40')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('41')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('41')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('42')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('42')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('43')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('43')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('44')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('44')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('45')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('45')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('46')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('46')->setWidth('5');

$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('47')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('47')->setWidth('30');

$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('48')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('48')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('49')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('49')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('50')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('50')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('51')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('51')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('52')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('52')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('53')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('53')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('54')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('54')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('55')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('55')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('56')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('56')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('57')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('57')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('58')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('58')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('59')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('59')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('60')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('60')->setWidth('5');

$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('61')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('61')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('62')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('62')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('63')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('63')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('64')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('64')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('65')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('65')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('66')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('66')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('67')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('67')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('68')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('68')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('69')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('69')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('70')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('70')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('71')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('71')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('72')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('72')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('73')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('73')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('74')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('74')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('75')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('75')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('76')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('76')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('77')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('77')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('78')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('78')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('79')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('79')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('80')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('80')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('81')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('81')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('82')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('82')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('83')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('83')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('84')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('84')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('85')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('85')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('86')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('86')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('87')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('87')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('88')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('88')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('89')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('89')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('90')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('90')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('91')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('91')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('92')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('92')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('93')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('93')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('94')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('94')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('95')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('95')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('96')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('96')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('97')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('97')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('98')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('98')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('99')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('99')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('100')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('100')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('101')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('101')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('102')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('102')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('103')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('103')->setWidth('5');


$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('106')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('106')->setWidth('15');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('107')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('107')->setWidth('30');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('108')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('108')->setWidth('20');

$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('109')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('109')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('110')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('110')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('111')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('111')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('112')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('112')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('113')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('113')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('114')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('114')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('115')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('115')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('116')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('116')->setWidth('5');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('117')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('117')->setWidth('5');

$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('118')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('118')->setWidth('30');

$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('119')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('119')->setWidth('6');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('120')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('120')->setWidth('6');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('121')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('121')->setWidth('6');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('122')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('122')->setWidth('6');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('123')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('123')->setWidth('6');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('124')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('124')->setWidth('6');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('125')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('125')->setWidth('6');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('126')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('126')->setWidth('6');
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('127')->setAutoSize(false);
$objPHPExcel->getActiveSheet()->getColumnDimensionByColumn('127')->setWidth('6');

 
// JUNTAMOS LAS CELDAS PARA LOS TITUTLOS
$objPHPExcel->setActiveSheetIndex(0)->mergeCells('H3:I3');
$objPHPExcel->getActiveSheet()->getStyle('H3')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('H3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('H3:I3')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('M3:X3');
$objPHPExcel->getActiveSheet()->getStyle('M3')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('M3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('M3:X3')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('Y3:AC3');
$objPHPExcel->getActiveSheet()->getStyle('Y3')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('Y3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('Y3:AC3')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('AD3:AH3');
$objPHPExcel->getActiveSheet()->getStyle('AD3')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('AD3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('AD3:AH3')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('AI3:AU3');
$objPHPExcel->getActiveSheet()->getStyle('AI3')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('AI3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('AI3:AU3')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('AW3:BI3');
$objPHPExcel->getActiveSheet()->getStyle('AW3')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('AW3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('AW3:BI3')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('BJ2:DB2');
$objPHPExcel->getActiveSheet()->getStyle('BJ2')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('BJ2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('BJ2:DB2')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('BJ3:BR3');
$objPHPExcel->getActiveSheet()->getStyle('BJ3')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('BJ3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('BJ3:BR3')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('BS3:CI3');
$objPHPExcel->getActiveSheet()->getStyle('BS3')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('BS3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('BS3:CI3')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('CJ3:CQ3');
$objPHPExcel->getActiveSheet()->getStyle('CJ3')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('CJ3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('CJ3:CQ3')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('CR3:CZ3');
$objPHPExcel->getActiveSheet()->getStyle('CR3')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('CR3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('CR3:CZ3')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('DA3:DB3');
$objPHPExcel->getActiveSheet()->getStyle('DA3')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('DA3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('DA3:DB3')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('DC2:EG2');
$objPHPExcel->getActiveSheet()->getStyle('DC2')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('DC2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('DC2:EG2')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('DF3:DH3');
$objPHPExcel->getActiveSheet()->getStyle('DF3')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('DF3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('DF3:DH3')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('DI3:DK3');
$objPHPExcel->getActiveSheet()->getStyle('DI3')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('DI3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('DI3:DK3')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('DL3:DN3');
$objPHPExcel->getActiveSheet()->getStyle('DL3')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('DL3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('DL3:DN3')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('DP3:DX3');
$objPHPExcel->getActiveSheet()->getStyle('DP3')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('DP3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('DP3:DX3')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('DY3:EG3');
$objPHPExcel->getActiveSheet()->getStyle('DY3')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('DY3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('DY3:EG3')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('EH2:EO2');
$objPHPExcel->getActiveSheet()->getStyle('EH2')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('EH2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('EH2:EO2')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('DA3:DB3');
$objPHPExcel->getActiveSheet()->getStyle('DA3:DB3')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('DA3:DB3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('DA3:DB3')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('EH3:EI3');
$objPHPExcel->getActiveSheet()->getStyle('EH3')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('EH3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('EH3:EI3')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);

$objPHPExcel->getActiveSheet()->getStyle('EJ3')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);

$objPHPExcel->setActiveSheetIndex(0)->mergeCells('EL3:EO3');
$objPHPExcel->getActiveSheet()->getStyle('EL3')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('EL3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('EL3:EO3')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);

//****************************************************************************************************************
//

$objPHPExcel->getActiveSheet()->getStyle('A4:EO4')->getAlignment()->setWrapText(true);
$objPHPExcel->getActiveSheet()->getStyle('A4:EO4')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
$objPHPExcel->getActiveSheet()->getStyle('A4:EO4')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);

/*************************************************LLENAMOS CON LA BASE DE DATOS*************************/

/********************************************************************************************************/



$fecha=date("Y");
$objPHPExcel->getActiveSheet()->setTitle($fecha);
$objPHPExcel->setActiveSheetIndex(0);


header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="monitoreo.xls"');
header('Cache-Control: max-age=0');

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');
exit;
?>