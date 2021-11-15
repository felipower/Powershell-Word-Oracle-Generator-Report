#
# @autor : Felipe Donoso Bastias, felipe@felipedonoso.cl
#
# @observacion: Debe estar instalados los componentes de SQL y WORD de powershell para que funcione
# Para ejecutar el script:
# powershell -executionpolicy bypass -file 01_FDB_Oracle_Word_Generator.ps1
#
#



#$encodage = New-Object System.Text.ASCIIEncoding 
#[System.Console]::OutputEncoding = $encodage

[Reflection.Assembly]::LoadWithPartialName("System.Data.OracleClient")

# Variables para el reporte word

#    <odc:ConnectionString>Provider=MSDAORA.1;Password=OCS.ora1234;User ID=ocs_m2c;Data Source=132.240.150.134:1521/dwqa.netlkjqsk.vcnqtlni.oraclevcn.com</odc:ConnectionString>

	# Esta variable cliente es mas bien la descripcion sera utilizada para poner nombre 
	$cliente="nombre_del_cliente"
	$dir=$(pwd)
	#$server="srvinfoaw"
	# Debe llevar "\" si es una instancia nombrada
	#	
	#$instancia="proprd.cl1.ocm.s1720959.oraclecloudatcustomer.com"
	#$instancia="dwprd.cl1.ocm.s1720959.oraclecloudatcustomer.com"
	#$instancia="biprd.cl1.ocm.s1720959.oraclecloudatcustomer.com"
	#$instancia="sapbiprd.cl1.ocm.s1720959.oraclecloudatcustomer.com"
	#$instancia="ctlmprd.cl1.ocm.s1720959.oraclecloudatcustomer.com"
	#esta es la misma de control M $instancia="arisdev.netlkjqsk.vcnqtlni.oraclevcn.com"
	#$instancia="dispfinprd.cl1.ocm.s1720959.oraclecloudatcustomer.com"
	#$instancia="ppffprd.cl1.ocm.s1720959.oraclecloudatcustomer.com"
	
	$instancia="CTMPRD_PDB1.netlkjqsk.vcnqtlni.oraclevcn.com"
	$ip="133.241.140.145"
	#$ip="10.108.3.6"
	$usuario="ocs_m2c"
	$clave="xxxxxx"
	$puerto="1521"
	$base=""
	$datasource=(echo $ip$instancia)
	# Timeout de la consulta Parametro Connect Timeout expresado en segundos (esto es el timeout solo para la query)
	$timeoutconsulta=60
	$headerdocumento="               Informe de status de plataforma de base de datos $server"
	$footerdocumento="IBM-GTS - Midrange DB"
	$titulodocumentoword="Reporte de status servidor $server"
	
	$time = (Get-Date).ToString("yyyyMMddHHm")
	$nombrearchivofinal="$dir\FDB_Oracle_Word_Generator_" + $time + "_" + $instancia + ".docx"
	
	# Edit here the name of template to use with own styles
	$archivotemplate="$dir\FDB_Oracle_Word_Generator__TEMPLATE03_NO_DELETE.docx"

	
# Copia del archivo de template
cp $archivotemplate $nombrearchivofinal

# No modificar esto
$Word = New-Object -ComObject Word.Application
# Esto es para que el documento se abra y veamos el word
#$Word.Visible = $True
#$Document = $Word.Documents.Add()
$Document = $Word.Documents.Open($nombrearchivofinal)
$Selection = $Word.Selection
$Selection.Font.Name="Arial"
$Selection.Font.Size=11

# Funciones necesarias para el reporte (creacion de las tablas de word)
Function Exportar-ResultadoQuery-A-TablaWord-Normal {
	Param (
			[parameter()]
			[string]$FormatoTabla ,
			[parameter()]
			[string]$QueryTexto
			
		)
   # Parametro Connect Timeout expresado en segundos (este es el timeout para la conexion a la instancia)
	#$connectionString = "Server=$dataSource;uid=$usuario; pwd=$clave;Database=$base;Integrated Security=False;Connect Timeout=300"

	#$connection = New-Object System.Data.SqlClient.SqlConnection
	$connection=New-Object DATA.OracleClient.OracleConnection("Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$ip)(PORT=$puerto))(CONNECT_DATA=(SERVICE_NAME=$instancia)));;User Id=$usuario;Password=$clave")
	#$connection.ConnectionString = $connectionString
	try{ 
		$connection.Open()
		$query = $QueryTexto
		$command = $connection.CreateCommand()
		$command.CommandTimeout=$timeoutconsulta
		$command.CommandText  =  $query
		echo "*** EJECUTANDO CONSULTA ==>  $query"
		 
			try{
				$result = $command.ExecuteReader()
				$table = new-object System.Data.DataTable
				$table.Load($result)
				$table | export-csv -Delimiter '|' -NoTypeInformation $dir\salida.csv
				
				#cat  $dir\salida.csv
				
				#If ((Get-Content "$dir\salida.csv") -eq $Null) {
				#	echo "N/A Sin resultados al respecto." > $dir\salida.csv
				#}
				
				#New-WordTable -Object $table -Columns 4 -Rows ($table.Count+1) –AsTable
			    
				#$Range = $Document.Range()
				$Range = @($Selection.Paragraphs)[-1].Range
				#$Range = @($Document.Paragraphs)[-1].Range
				
				$text=(cat $dir\salida.csv | Out-String) -replace '"',''
				#echo $text
				$Range.Text = "$text"
				#$separator=[Microsoft.Office.Interop.Word.WdTableFieldSeparator]::wdSeparateByCommas
				$Separator="|"
				#$table=$Range.ConvertToTable Separator:=$Separator , Format:=wdTableFormatGrid4  
				$table=$Range.ConvertToTable($Separator)
				
				# Ojo con esto es para que se repitan los encabezados
				$table.Rows.item(1).Headingformat=-1

				# OJO AQUI HAY UN BUG CADA VEZ QUE SE LE DA FORMATO A UNA TABLA EXISTENTE
				# LAS COLUMNAS SE DESORDENAN
				# El formato de las tablas puede ser consultado en el link:
				# https://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdtableformat%28v=office.14%29.aspx
				# Aqui se le pone el formato a las tablas por ahora se esta ocupando el segundo metodo
				#$table.AutoFormat([Microsoft.Office.Interop.Word.WdTableFormat]::$FormatoTabla)
				#$table.AutoFormat([Microsoft.Office.Interop.Word.WdTableFormat]::wdTableFormatElegant)
				#$table.AutoFormat([Microsoft.Office.Interop.Word.WdTableFormat]::wdTableFormatSimple3)
				#$table.AutoFormat([Microsoft.Office.Interop.Word.WdTableFormat]::wdTableFormatGrid4)
				
				
				
				#$table.Style = "Medium Shading 1 - Accent 1";
				#table.AutoFormat([Microsoft.Office.Interop.Word.WdTableFormatApply]::wdTableFormatApplyHeadingRows)
			

				#$table.AutoFit = $true
				#$table.AutoFormat([Microsoft.Office.Interop.Word.WdTableFormat]::wdAutoFitWindow)
				$table.Range.Font.Size = 8
				#$Selection.Style = 'normal'
				$table.Range.Style = 'FDB_TABLA_REPORTE'
				#$table.Range.paragraphFormat.Alignment = 1
				
				# se desordena el grafico hay un error con el autoformato de tablas para esta version de powershell
				#$table.AllowAutofit = $true
				$table.ApplyStyleHeadingRows = $true
				
				
				
				#$table.Range.paragraphFormat.Alignment = wdAlignParagraphCenter
				#$table.AutoFormat(19)
				
				#$table.Rows.HeadingFormat = True
				#$table.Rows.HeadingFormat = -1
				#$table.Range.Orientation = 1
				#$table.AutoFormat([Microsoft.Office.Interop.Word.AutoFitBehavior]::wdAutoFitWindow) 
				#$table.Range.Style = "Table Grid"
				$table.Borders.InsideLineStyle = 1
				$table.Borders.OutsideLineStyle = 1
				#$Selection.PageSetup.Orientation = 1
				#$table.Rows.HeadingFormat = True

			}
			catch{
				echo "hubo un problema en la ejecucion de la consulta anterior (revisar si hay timeout en la conexion)"
				echo $query
				$connection.Open()
				
				
			}
	}
	catch{
		echo $connectionString
		echo "No fue posible conectar con Servicio de BD $server$instancia (Timeout)" 
	}
	$connection.Close()

}







Function Exportar-ResultadoQuery-A-Grafico-1-Columna-Normal {
	Param (
			[parameter()]
			[string]$FormatoGrafico,
			[parameter()]
			[string]$RutaImagen,
			[parameter()]
			[string]$TituloGrafico,
			[parameter()]
			[string]$QueryTexto
		)
   # Parametro Connect Timeout expresado en segundos (este es el timeout para la conexion a la instancia)
	$connectionString = "Server=$dataSource;uid=$usuario; pwd=$clave;Database=$base;Integrated Security=False;Connect Timeout=300"

	$connection = New-Object System.Data.SqlClient.SqlConnection
	$connection.ConnectionString = $connectionString
	try{ 
		$connection.Open()
		$query = $QueryTexto
		$command = $connection.CreateCommand()
		$command.CommandTimeout=$timeoutconsulta
		echo $query
		$command.CommandText  =  $query
		#echo $query
			try{
				$result = $command.ExecuteReader()
				$table = new-object System.Data.DataTable
				$table.Load($result)
				$table | export-csv -Delimiter ';' -NoTypeInformation $dir\salida.csv
				
				[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")

				# chart object
				   $chart1 = New-object System.Windows.Forms.DataVisualization.Charting.Chart
				   $chart1.Width = 600
				   $chart1.Height = 600
				   $chart1.BackColor = [System.Drawing.Color]::White

				# title 
				   [void]$chart1.Titles.Add("$TituloGrafico")
				   $chart1.Titles[0].Font = "Arial,13pt"
				   $chart1.Titles[0].Alignment = "topLeft"

				# chart area 
				   $chartarea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
				   $chartarea.Name = "ChartArea1"
				   $chartarea.AxisY.Title = "titulo eje y"
				   $chartarea.AxisX.Title = "titulo eje x"
				   $chartarea.AxisY.Interval = 100
				   $chartarea.AxisX.Interval = 1
				   $chart1.ChartAreas.Add($chartarea)

				# legend 
				   $legend = New-Object system.Windows.Forms.DataVisualization.Charting.Legend
				   $legend.name = "Legend1"
				   $chart1.Legends.Add($legend)

				# Procedemos a llenar el datasource
					$result = $command.ExecuteReader()
					$table = new-object System.Data.DataTable
					$table.Load($result)
					$table | export-csv -Delimiter ';' -NoTypeInformation $dir\salida_grafico.csv
				    #$datasource_grafico = Get-Process | sort PrivateMemorySize -Descending  | Select-Object -First 5
					$datasource_grafico = ((cat $dir\salida_grafico.csv | Out-String) -replace ';','     ') -replace '"',''
					
					$nombre_serie_1=((cat $dir\salida_grafico.csv |  select -first 1 | Out-String) -replace '"','')  | %{ $_.Split(';')[1]; }
					
					
				# data series
				   [void]$chart1.Series.Add($nombre_serie_1)
				   $chart1.Series["VirtualMem"].ChartType = "Column"
				   $chart1.Series["VirtualMem"].BorderWidth  = 3
				   $chart1.Series["VirtualMem"].IsVisibleInLegend = $true
				   $chart1.Series["VirtualMem"].chartarea = "ChartArea1"
				   $chart1.Series["VirtualMem"].Legend = "Legend1"
				   $chart1.Series["VirtualMem"].color = "#62B5CC"
				   $datasource_grafico | ForEach-Object {$chart1.Series[$nombre_serie_1].Points.addxy( $_.Name , ($_.pctused)) }

				# data series 2
				#   [void]$chart1.Series.Add("PrivateMem")
				#   $chart1.Series["PrivateMem"].ChartType = "Column"
				#  $chart1.Series["PrivateMem"].IsVisibleInLegend = $true
				#   $chart1.Series["PrivateMem"].BorderWidth  = 3
				#   $chart1.Series["PrivateMem"].chartarea = "ChartArea1"
				#   $chart1.Series["PrivateMem"].Legend = "Legend1"
				#   $chart1.Series["PrivateMem"].color = "#E3B64C"
				#   $datasource | ForEach-Object {$chart1.Series["PrivateMem"].Points.addxy( $_.Name , ($_.PrivateMemorySize / 1000000)) }

				# save chart
				   $chart1.SaveImage("$RutaImagen","png")

			}
			catch{
				echo "hubo un problema en la ejecucion de la consulta anterior (revisar si hay timeout en la conexion)"
				echo $query
				$connection.Open()
				
			}
	}
	catch{
		echo $connectionString
		echo "$server$instancia ;<font color=FF0000><b>Revisar Instancia</b></font>;NA;NA;NA;NA;NA;NA" 
	}
	$connection.Close()
}




# Se comienza la generacion del reporte





#
#para ver el listado de fuentes 
#[Enum]::GetNames([Microsoft.Office.Interop.Word.WdBuiltinStyle]) | ForEach {
#    [pscustomobject]@{Style=$_}
#} | Format-Wide -Property Style -Column 4


###############################################################################
#
# Esto es para el encabezado y footer
# 
# 

#
# Esto es para automatizar el encabezado
# 
# $Section = $Document.Sections.Item(1);
# $Header = $Section.Headers.Item(1);
# $Footer = $Section.Footers.Item(1);
# $Header.Range.Font.Name="Calibri"
# $Header.Range.Font.Size = 10
# $Header.Range.Font.Bold = $true
# $Header.Range.Text = "$headerdocumento";
# $Header.Range.InlineShapes.AddPicture("$dir\header.png")
#$Footer.Range.Font.Name="Calibri"
#$Footer.Range.Font.Italic=$True
#$Footer.Range.Font.Size = 10
#$Footer.Range.Font.Bold = $true
#$Footer.Range.Text = $footerdocumento;
#$Footer.Range.paragraphFormat.Alignment = 'wdAlignParagraphCenter'

# Esto es para agregar el  numero de pagina a los footers del documento
#$c = $Document.Sections.Item(1).Footers.Item(1).PageNumbers.Add($wdAlignPageNumberRight) 

# Create a Table of Contents (ToC)
# No borrar esta parte pues es la tabla de contenido
#$Toc = $Document.TablesOfContents.Add($Section.Range);


$Word.Selection.EndKey(6,0)

#$Selection.Style = 'Title'
#$Selection.TypeText($titulodocumentoword)
#$Selection.TypeParagraph()


###############################################################################
#
# Aqui comienz el cuerpo del documento 
#
#

#$BuiltinProperties = $document.BuiltInDocumentProperties

# Esto es para agregar titulo al documento

    #$Selection.GoTo(1,2,$null,1)
    #Insert cover page
    #$Selection.InsertFile("$dir\portada.docx")


	
# Acentos
# á $Selection.insertsymbol(225)
# é $Selection.insertsymbol(233)
# í $Selection.insertsymbol(237)
# ó $Selection.insertsymbol(243)
# ú $Selection.insertsymbol(250)

# Saltos de linea dentro de un TypeText o de un echo
# `n

# Cambiar de Orientacion la pagina
#$Word.Selection.PageSetup.Orientation = 1
#$Document.PageSetup.Orientation = 1




$Selection.InsertBreak()
#$Word.Selection.PageSetup.Orientation = 0
$Selection.Style = 'HD1'
$Selection.TypeText("Oracle Database " + $instancia)
$Selection.TypeParagraph()
$Selection.Style = 'normal'
$Selection.TypeText("`nThis chapter will show information about this PDB $instancia and all their configurations.`n")
$Selection.TypeParagraph()


#$Selection.InsertBreak()
$Selection.TypeText("`n`n")
$Selection.Style = 'HD2'
$Selection.TypeText("PDBs")
$Selection.TypeParagraph()
$Selection.Style = 'normal'
$Selection.TypeText("The next Table show information about the PDBs that is created on Container.`n")
$Selection.TypeParagraph()
# Se sugiere usar single quote para la query y no double quote
# De esa manera en las querys se podra utilizar el set caracter dolar
# pues muchas consultas de oracle las tienen
Exportar-ResultadoQuery-A-TablaWord-Normal  -FormatoTabla "wdTableFormatList4" -QueryTexto '
select PDB_ID,PDB_NAME,STATUS,LOGGING,FORCE_LOGGING,FORCE_NOLOGGING,CREATION_TIME from CDB_PDBS
'
$Word.Selection.EndKey(6,0)



#$Selection.InsertBreak()
$Selection.TypeText("`n`n")
$Selection.Style = 'HD2'
$Selection.TypeText("Service Names PDBs")
$Selection.TypeParagraph()
$Selection.Style = 'normal'
$Selection.TypeText("The next Table show information about the service names created on every PDBs.`n")
$Selection.TypeParagraph()
Exportar-ResultadoQuery-A-TablaWord-Normal  -FormatoTabla "wdTableFormatList4" -QueryTexto '
select service_id,name,creation_date,pdb from CDB_SERVICES 
'
$Word.Selection.EndKey(6,0)



#$Selection.InsertBreak()
$Selection.TypeText("`n`n")
$Selection.Style = 'HD2'
$Selection.TypeText("ASM Disk Groups")
$Selection.TypeParagraph()
$Selection.Style = 'normal'
$Selection.TypeText("The following table shows the different ASM disk groups that are available to be occupied by the Database Instances.`n")
$Selection.TypeParagraph()
Exportar-ResultadoQuery-A-TablaWord-Normal  -FormatoTabla "wdTableFormatList4" -QueryTexto '
SELECT
    name                                     group_name
  , allocation_unit_size                     allocation_unit_size
  , type                                     type
  , round(total_mb/1024,2)                                 total_gb
  , round((total_mb - free_mb)/1024,2)                     used_gb
  , ROUND((1- (free_mb / greatest(total_mb,1)))*100, 2)  pct_used
FROM    v$asm_diskgroup
ORDER BY    6 desc
'
$Word.Selection.EndKey(6,0)


#$Selection.InsertBreak()
$Selection.TypeText("`n`n")
$Selection.Style = 'HD2'
$Selection.TypeText("Database Instances")
$Selection.TypeParagraph()
$Selection.Style = 'normal'
$Selection.TypeText("The following Information corresponds to each Database Instances along with their corresponding version.`n")
$Selection.TypeParagraph()
Exportar-ResultadoQuery-A-TablaWord-Normal  -FormatoTabla "wdTableFormatList4" -QueryTexto '
select inst_id,INSTANCE_name, host_name,version, status,archiver,thread# from GV$instance
'
$Word.Selection.EndKey(6,0)



# Para esta query conviene mejor poner un salto de linea porque son muchos datos
$Selection.InsertBreak()
$Selection.TypeText("`n`n")
$Selection.Style = 'HD2'
$Selection.TypeText("Database parameters (Non Default)")
$Selection.TypeParagraph()
$Selection.Style = 'normal'
$Selection.TypeText("The following Table are the non-default database parameters of each Database Instance. Only the parameters that have a different value than default of the instance are attached in order not to provide a too exhaustive list.`n")
$Selection.TypeParagraph()
Exportar-ResultadoQuery-A-TablaWord-Normal  -FormatoTabla "wdTableFormatList4" -QueryTexto '
select * from (
select inst_id, name "Parametro/INST_ID ==>", value from GV$parameter p
where isdefault = ''FALSE''
order by name,inst_id
)
pivot 
(
min (value)
-- aumentar el numero de instancias si es que son mas
for (inst_id) in (1,2)
) order by 1
'
$Word.Selection.EndKey(6,0)



# Para esta query conviene mejor poner un salto de linea porque son muchos datos
$Selection.InsertBreak()
$Selection.TypeText("`n`n")
$Selection.Style = 'HD2'
$Selection.TypeText("Memory Distribution")
$Selection.TypeParagraph()
$Selection.Style = 'normal'
$Selection.TypeText("The following information shows the current values that are being used for the memory distribution of the database instances. Here you can see the values assigned to components of the SGA and PGA.`n")
$Selection.TypeParagraph()
Exportar-ResultadoQuery-A-TablaWord-Normal  -FormatoTabla "wdTableFormatList4" -QueryTexto '
select (select instance_name from GV$instance where m.inst_id = inst_id ) instance
, component,parameter,round(final_size/1024/1024,2) SIZE_MB
 from GV$MEMORY_RESIZE_OPS m
--where COMPONENT like ''%SGA%''
where initial_size = 0
--and component != ''ASM Buffer Cache''
order by component, 1
'
$Word.Selection.EndKey(6,0)



$Selection.InsertBreak()
$Selection.TypeText("`n`n")
$Selection.Style = 'HD2'
$Selection.TypeText("Database Controlfile")
$Selection.TypeParagraph()
$Selection.Style = 'normal'
$Selection.TypeText("The following files correspond to the locations of the current database controls.`n")
$Selection.TypeParagraph()
# Aqui en esta parte al momento de ejecuta la query
# y escribir el resultado en el archivo word se cae
# al parecer hay un problema o error al leer la columna status de la V$controlfile
Exportar-ResultadoQuery-A-TablaWord-Normal  -FormatoTabla "wdTableFormatList4" -QueryTexto '
select name,is_recovery_dest_file,block_size,file_size_blks from V$controlfile
'
$Word.Selection.EndKey(6,0)
# select * from V$CONTROLFILE


#$Selection.InsertBreak()
$Selection.TypeText("`n`n")
$Selection.Style = 'HD2'
$Selection.TypeText("Redolog Files")
$Selection.TypeParagraph()
$Selection.Style = 'normal'
$Selection.TypeText("The following information shows the redolog members of each group and thread (Instance) along with the size in MB of each file.`n")
$Selection.TypeParagraph()
Exportar-ResultadoQuery-A-TablaWord-Normal  -FormatoTabla "wdTableFormatList4" -QueryTexto '
SELECT a.GROUP#, a.THREAD#
--, a.SEQUENCE#,
--a.ARCHIVED, a.STATUS
 , b.MEMBER AS REDOLOG_FILE_NAME,
--b.type,
 (a.BYTES/1024/1024) AS SIZE_MB FROM v$log a
JOIN v$logfile b ON a.Group#=b.Group#
ORDER BY a.GROUP#
'
$Word.Selection.EndKey(6,0)


$Selection.InsertBreak()
$Selection.TypeText("`n`n")
$Selection.Style = 'HD2'
$Selection.TypeText("Database Patch/PSU ")
$Selection.TypeParagraph()
$Selection.Style = 'normal'
$Selection.TypeText("The table below shows the different patches or updates (PSU or FIX) applied in the database. The first column indicates the date of application of these.`n")
$Selection.TypeParagraph()
Exportar-ResultadoQuery-A-TablaWord-Normal  -FormatoTabla "wdTableFormatList4" -QueryTexto '
select patch_id,patch_uid,action,status,action_time,bundle_series,bundle_id, description,con_id from CDB_REGISTRY_SQLPATCH
'
$Word.Selection.EndKey(6,0)


$Selection.InsertBreak()
$Selection.TypeText("`n`n")
$Selection.Style = 'HD2'
$Selection.TypeText("Tablespaces")
$Selection.TypeParagraph()
$Selection.Style = 'normal'
$Selection.TypeText("The following table describes the list of database tablespaces along with their current sizes (Only PDB).`n")
$Selection.TypeParagraph()
Exportar-ResultadoQuery-A-TablaWord-Normal  -FormatoTabla "wdTableFormatList4" -QueryTexto '
select tablespace_name, block_size
,status, contents, extent_management,segment_space_management,bigfile--,encrypted
from dba_tablespaces
'
$Word.Selection.EndKey(6,0)


$Selection.InsertBreak()
$Selection.TypeText("`n`n")
$Selection.TypeParagraph()
$Selection.Style = 'normal'
$Selection.TypeText("The current sizes of data tablespaces are as follows (expressed in MB).`n")
$Selection.TypeParagraph()
Exportar-ResultadoQuery-A-TablaWord-Normal  -FormatoTabla "wdTableFormatList4" -QueryTexto '
SELECT   
tm.tbs Tablespace,
           to_char((tm.mb - free.mb),''999999999999990D00'') UsedMB,
         to_char(free.mb,''999999999999990D00'') FreeMB,
           to_char(tm.mb,''999999999999990D00'') TotalMB,
         to_char(((tm.mb - free.mb) / tm.mb) * 100,''990D00'') "pct. Used"
    FROM (SELECT   tablespace_name tbs, SUM (BYTES) / 1024 / 1024 mb
              FROM dba_data_files GROUP BY tablespace_name) tm,
         (SELECT   tablespace_name tbs, SUM (BYTES) / 1024 / 1024 mb
              FROM dba_free_space GROUP BY tablespace_name  ) free
   WHERE tm.tbs = free.tbs(+)
ORDER BY 5 DESC
'
$Word.Selection.EndKey(6,0)




$Selection.InsertBreak()
$Selection.TypeText("`n`n")
$Selection.Style = 'HD2'
$Selection.TypeText("Datafiles and Tempfiles")
$Selection.TypeParagraph()
$Selection.Style = 'normal'
$Selection.TypeText("The list of datafiles and tempfiles of each database tablespaces is attached. The following table shows the database files (only production PDB):`n")
$Selection.TypeParagraph()
Exportar-ResultadoQuery-A-TablaWord-Normal  -FormatoTabla "wdTableFormatList4" -QueryTexto '
select file_name,tablespace_name,round(bytes/1024/1024,2) SIZE_MB, autoextensible,increment_by  from dba_data_files
'
$Word.Selection.EndKey(6,0)

#$Selection.InsertBreak()
$Selection.TypeText("`n`n")
$Selection.TypeParagraph()
$Selection.Style = 'normal'
$Selection.TypeText("Regarding to tempfiles, the table below shows the database tempfiles:`n")
$Selection.TypeParagraph()
Exportar-ResultadoQuery-A-TablaWord-Normal  -FormatoTabla "wdTableFormatList4" -QueryTexto '
select file_name,tablespace_name,round(bytes/1024/1024,2) SIZE_MB, autoextensible,increment_by  from dba_temp_files
'
$Word.Selection.EndKey(6,0)




$Selection.InsertBreak()
$Selection.TypeText("`n`n")
$Selection.Style = 'HD2'
$Selection.TypeText("NLS Character Set")
$Selection.TypeParagraph()
$Selection.Style = 'normal'
$Selection.TypeText("The information of the NLS parameters of the database is attached in the following table:`n")
$Selection.TypeParagraph()
Exportar-ResultadoQuery-A-TablaWord-Normal  -FormatoTabla "wdTableFormatList4" -QueryTexto '
select * from nls_database_parameters
'
$Word.Selection.EndKey(6,0)


$Selection.InsertBreak()
$Selection.TypeText("`n`n")
$Selection.Style = 'HD2'
$Selection.TypeText("Users")
$Selection.TypeParagraph()
$Selection.Style = 'normal'
$Selection.TypeText("The user list created on the Database (oracle_maintained = ''N''):`n")
$Selection.TypeParagraph()
Exportar-ResultadoQuery-A-TablaWord-Normal  -FormatoTabla "wdTableFormatList4" -QueryTexto '
select username,default_tablespace,temporary_tablespace,created,profile from dba_users
where oracle_maintained = ''N''
'
$Word.Selection.EndKey(6,0)




$Selection.InsertBreak()
$Selection.TypeText("`n`n")
$Selection.Style = 'HD2'
$Selection.TypeText("Access Control List ACL ")
$Selection.TypeParagraph()
$Selection.Style = 'normal'
$Selection.TypeText("The next information is regarding the ACL created on this database (DBA_HOST_ACLS only this PDB):`n")
$Selection.TypeParagraph()
Exportar-ResultadoQuery-A-TablaWord-Normal  -FormatoTabla "wdTableFormatList4" -QueryTexto '
select HOST	, LOWER_PORT	, UPPER_PORT, 	ACL	,	ACL_OWNER 
from DBA_HOST_ACLS
'
$Word.Selection.EndKey(6,0)

#$Selection.InsertBreak()
$Selection.TypeText("`n`n")
$Selection.TypeParagraph()
$Selection.Style = 'normal'
$Selection.TypeText("The next information is regarding the ACL privileges created on this database (DBA_NETWORK_ACL_PRIVILEGES only this PDB):`n:`n")
$Selection.TypeParagraph()
Exportar-ResultadoQuery-A-TablaWord-Normal  -FormatoTabla "wdTableFormatList4" -QueryTexto '
select ACL	,	PRINCIPAL	,PRIVILEGE	,IS_GRANT	,INVERT	,START_DATE	,END_DATE	,ACL_OWNER 
from DBA_NETWORK_ACL_PRIVILEGES
'
$Word.Selection.EndKey(6,0)

$Selection.InsertBreak()
$Selection.TypeText("`n`n")
$Selection.Style = 'HD2'
$Selection.TypeText("Connection String tnsnames.ora")
$Selection.TypeParagraph()
$Selection.Style = 'normal'
$Selection.TypeText("The following string is the one used to connect to the database. This is the same string that exists in the tnsnames.ora file:`n")
$Selection.TypeParagraph()
#Exportar-ResultadoQuery-A-TablaWord-Normal  -FormatoTabla "wdTableFormatList4" -QueryTexto '
#select * from nls_database_parameters
#'
$Word.Selection.EndKey(6,0)








$Document.SaveAs([ref]$nombrearchivofinal,[ref]$SaveFormat::wdFormatDocument)

#Out-File -FilePath $nombrearchivofinal -Encoding UTF8 -Force

$word.Quit()



#
#Esto es para liberar la memoria es importante
#
$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$word)
[gc]::Collect()
[gc]::WaitForPendingFinalizers()
Remove-Variable Word 

