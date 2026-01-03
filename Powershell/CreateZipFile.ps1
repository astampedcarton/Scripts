Clear

#Update these as needed
$ziploc = "c:\temp\p21.zip"
$root = "G:\My Drive\Pharma_Life\CDISC"

$rgsrc = "$root\Define\DefineV215\examples\Define-XML-2-1-ADaM\adam\adrg.pdf"
$adamdata = "$root\pilot_data\ADaM\*.xpt"
$adamdefine = "$root\pilot_data\ADaM\define.xml"
$definestyle = "$root\pilot_data\ADaM\define.xsl"

# SDTM data to be copied
$sdtmdm = "$root\pilot_data\SDTM\dm.xpt"
$sdtmae = "$root\pilot_data\SDTM\ae.xpt"
$sdtmex = "$root\pilot_data\SDTM\ex.xpt"
	

#Add all the files that needs to be added to the zip
$file4Zip = @(
  $rgsrc,
  $adamdefine,
  $definestyle,
  $adamdata,
  $sdtmdm,
  $sdtmae,
  $sdtmex
)

#Zip the sources into a single file
Compress-Archive -Path $file4Zip -DestinationPath $ziploc


#open the location of the zip file
$loc = Split-Path $ziploc

Start-Process $loc
