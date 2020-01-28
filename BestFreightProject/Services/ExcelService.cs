using AutoMapper;
using BestFreightProject.Dtos;
using BestFreightProject.Entities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace BestFreightProject.Services
{
    public class ExcelService : IExcelService
    {
        private IMapper mapper { get; }
        public ExcelService(IMapper mapper) => this.mapper = mapper;
        public void CreateExcelFile(ExcelCreateDto excelCreate)
        {
            var excel = mapper.Map<Excel>(excelCreate);
            var excelString = HtmlExcelDocument(excel);
        }

        public FileStream GetExcel(string path)
        {
            throw new NotImplementedException();
        }

        private string HtmlExcelDocument(Excel excel)
        {
            var excelString = @"<html xmlns:o=""urn: schemas - microsoft - com:office: office""
xmlns: x = ""urn:schemas-microsoft-com:office:excel""
xmlns = ""http://www.w3.org/TR/REC-html40"" >

< head >
< meta http - equiv = Content - Type content = ""text/html; charset=windows-1252"" >
       < meta name = ProgId content = Excel.Sheet >
          < meta name = Generator content = ""Microsoft Excel 15"" >
             < link rel = File - List href = ""Solicitud%20de%20Cotización_archivos/filelist.xml"" >
                  < style id = ""Solicitud de Cotización_3834_Styles"" >
                   < !--table

     {
                mso - displayed - decimal - separator:""\."";
                mso - displayed - thousand - separator:""\,"";
            }
.xl153834

    {
                padding - top:1px;
                padding - right:1px;
                padding - left:1px;
                mso - ignore:padding;
            color: black;
                font - size:11.0pt;
                font - weight:400;
                font - style:normal;
                text - decoration:none;
                font - family:Calibri;
                mso - generic - font - family:auto;
                mso - font - charset:0;
                mso - number - format:General;
                text - align:general;
                vertical - align:bottom;
                mso - background - source:auto;
                mso - pattern:auto;
                white - space:nowrap;
            }
.xl633834

    {
                padding - top:1px;
                padding - right:1px;
                padding - left:1px;
                mso - ignore:padding;
            color: white;
                font - size:12.0pt;
                font - weight:700;
                font - style:normal;
                text - decoration:none;
                font - family:Calibri;
                mso - generic - font - family:auto;
                mso - font - charset:0;
                mso - number - format:General;
                text - align:center;
                vertical - align:middle;
            border: .5pt solid black;
            background:#23435F;
	mso - pattern:black none;
                white - space:normal;
            }
.xl643834

    {
                padding - top:1px;
                padding - right:1px;
                padding - left:1px;
                mso - ignore:padding;
            color: black;
                font - size:12.0pt;
                font - weight:700;
                font - style:normal;
                text - decoration:none;
                font - family:Calibri;
                mso - generic - font - family:auto;
                mso - font - charset:0;
                mso - number - format:General;
                text - align:center;
                vertical - align:middle;
            border: .5pt solid black;
            background: white;
                mso - pattern:black none;
                white - space:normal;
            }
.xl653834

    {
                padding - top:1px;
                padding - right:1px;
                padding - left:1px;
                mso - ignore:padding;
            color: white;
                font - size:12.0pt;
                font - weight:400;
                font - style:normal;
                text - decoration:none;
                font - family:Calibri;
                mso - generic - font - family:auto;
                mso - font - charset:0;
                mso - number - format:General;
                text - align:general;
                vertical - align:bottom;
            border: .5pt solid black;
            background:#23435F;
	mso - pattern:black none;
                white - space:nowrap;
            }
.xl663834

    {
                padding - top:1px;
                padding - right:1px;
                padding - left:1px;
                mso - ignore:padding;
            color: white;
                font - size:16.0pt;
                font - weight:700;
                font - style:normal;
                text - decoration:none;
                font - family:Verdana;
                mso - generic - font - family:auto;
                mso - font - charset:0;
                mso - number - format:General;
                text - align:center;
                vertical - align:middle;
            background:#220835;
	mso - pattern:black none;
                white - space:normal;
            }
.xl673834

    {
                padding - top:1px;
                padding - right:1px;
                padding - left:1px;
                mso - ignore:padding;
            color: black;
                font - size:12.0pt;
                font - weight:700;
                font - style:normal;
                text - decoration:none;
                font - family:Calibri;
                mso - generic - font - family:auto;
                mso - font - charset:0;
                mso - number - format:General;
                text - align:general;
                vertical - align:bottom;
                mso - background - source:auto;
                mso - pattern:auto;
                white - space:nowrap;
            }
.xl683834

    {
                padding - top:1px;
                padding - right:1px;
                padding - left:1px;
                mso - ignore:padding;
            color:#F3A42B;
	font - size:12.0pt;
                font - weight:700;
                font - style:normal;
                text - decoration:none;
                font - family:Calibri;
                mso - generic - font - family:auto;
                mso - font - charset:0;
                mso - number - format:General;
                text - align:center;
                vertical - align:middle;
            border: .5pt solid black;
            background: white;
                mso - pattern:black none;
                white - space:normal;
            }
.xl693834

    {
                padding - top:1px;
                padding - right:1px;
                padding - left:1px;
                mso - ignore:padding;
            color: white;
                font - size:12.0pt;
                font - weight:400;
                font - style:normal;
                text - decoration:none;
                font - family:Calibri;
                mso - generic - font - family:auto;
                mso - font - charset:0;
                mso - number - format:General;
                text - align:center;
                vertical - align:middle;
            border: .5pt solid black;
            background:#23435F;
	mso - pattern:black none;
                white - space:normal;
            }
.xl703834

    {
                padding - top:1px;
                padding - right:1px;
                padding - left:1px;
                mso - ignore:padding;
            color:#F3A42B;
	font - size:12.0pt;
                font - weight:700;
                font - style:normal;
                text - decoration:none;
                font - family:Calibri;
                mso - generic - font - family:auto;
                mso - font - charset:0;
                mso - number - format:General;
                text - align:general;
                vertical - align:bottom;
            border: .5pt solid black;
            background:#23435F;
	mso - pattern:black none;
                white - space:nowrap;
            }
            -->
            </ style >
            </ head >
            

            < body >
            < !--[if !excel]> &nbsp; &nbsp;< ![endif]-- >
               < !--La siguiente información se generó mediante el Asistente para publicar como
página web de Microsoft Excel.-- >
< !--Si se vuelve a publicar el mismo elemento desde Excel, se reemplazará toda
 la información comprendida entre las etiquetas DIV.-->
 < !----------------------------->
 < !--INICIO DE LOS RESULTADOS DEL ASISTENTE PARA PUBLICAR COMO PÁGINA WEB DE
  EXCEL -->
  < !----------------------------->
  

  < div id = ""Solicitud de Cotización_3834"" align = center x: publishsource = ""Excel"" >
        

        < table border = 0 cellpadding = 0 cellspacing = 0 width = 1658 style = 'border-collapse:
 collapse; table - layout:fixed; width: 1243pt'>
      < col width = 64 style = 'width:48pt' >
        
         < col width = 169 style = 'mso-width-source:userset;mso-width-alt:6180;width:127pt' >
           
            < col width = 64 style = 'width:48pt' >
              
               < col width = 63 style = 'mso-width-source:userset;mso-width-alt:2304;width:47pt' >
                 
                  < col width = 119 style = 'mso-width-source:userset;mso-width-alt:4352;width:89pt' >
                    
                     < col width = 64 style = 'width:48pt' >
                       
                        < col width = 169 style = 'mso-width-source:userset;mso-width-alt:6180;width:127pt' >
                          
                           < col width = 64 style = 'width:48pt' >
                             
                              < col width = 119 style = 'mso-width-source:userset;mso-width-alt:4352;width:89pt' >
                                
                                 < col width = 64 style = 'width:48pt' >
                                   
                                    < col width = 119 style = 'mso-width-source:userset;mso-width-alt:4352;width:89pt' >
                                      
                                       < col width = 214 style = 'mso-width-source:userset;mso-width-alt:7826;width:161pt' >
                                         
                                          < col width = 119 style = 'mso-width-source:userset;mso-width-alt:4352;width:89pt' >
                                            
                                             < col width = 64 style = 'width:48pt' >
                                               
                                                < col width = 119 style = 'mso-width-source:userset;mso-width-alt:4352;width:89pt' >
                                                  
                                                   < col width = 64 style = 'width:48pt' >
                                                     
                                                      < tr height = 20 style = 'height:15.0pt' >
                                                        
                                                          < td height = 20 class=xl153834 width = 64 style='height:15.0pt;width:48pt'></td>
  <td class=xl153834 width = 169 style='width:127pt'></td>
  <td class=xl153834 width = 64 style='width:48pt'></td>
  <td class=xl153834 width = 63 style='width:47pt'></td>
  <td class=xl153834 width = 119 style='width:89pt'></td>
  <td class=xl153834 width = 64 style='width:48pt'></td>
  <td class=xl153834 width = 169 style='width:127pt'></td>
  <td class=xl153834 width = 64 style='width:48pt'></td>
  <td class=xl153834 width = 119 style='width:89pt'></td>
  <td class=xl153834 width = 64 style='width:48pt'></td>
  <td class=xl153834 width = 119 style='width:89pt'></td>
  <td class=xl153834 width = 214 style='width:161pt'></td>
  <td class=xl153834 width = 119 style='width:89pt'></td>
  <td class=xl153834 width = 64 style='width:48pt'></td>
  <td class=xl153834 width = 119 style='width:89pt'></td>
  <td class=xl153834 width = 64 style='width:48pt'></td>
 </tr>
 <tr height = 20 style='height:15.0pt'>
  <td height = 20 class=xl153834 style = 'height:15.0pt' ></ td >
  
    < td colspan=15 rowspan=2 class=xl663834 width = 1594 style='width:1195pt'>www.BestFreightSearch.com
  –<span style = 'mso-spacerun:yes' > </ span > The Best Freight Searcher in the
        World</td>
 </tr>
 <tr height = 20 style= 'height:15.0pt' >
      
        < td height= 20 class=xl153834 style = 'height:15.0pt' ></ td >
        
         </ tr >
        
         < tr height=21 style='height:15.75pt'>
  <td height = 21 class=xl153834 style = 'height:15.75pt' ></ td >
  
    < td colspan=6 class=xl673834>%%FreightType%%</td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
 </tr>
 <tr height = 21 style='height:15.75pt'>
  <td height = 21 class=xl153834 style = 'height:15.75pt' ></ td >
  
    < td colspan=6 class=xl673834>%%SubFreightType%%</td>
  <td class=xl153834>%%QuotationNumber%%/td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td colspan = 4 class=xl153834>%%QuotationDate%%</td>
 </tr>
 <tr height = 20 style='height:15.0pt'>
  <td height = 20 class=xl153834 style = 'height:15.0pt' ></ td >
  
    < td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
 </tr>
 <tr height = 21 style='height:15.75pt'>
  <td height = 21 class=xl153834 style = 'height:15.75pt' ></ td >
  
    < td colspan=3 class=xl633834 width = 296 style='width:222pt'>Name of Company</td>
  <td colspan = 3 class=xl633834 width = 352 style='border-left:none;width:264pt'>Contact
      Person</td>
  <td colspan = 3 class=xl633834 width = 247 style='border-left:none;width:185pt'>E-mail</td>
  <td colspan = 3 class=xl633834 width = 452 style='border-left:none;width:339pt'>Telephone</td>
  <td colspan = 3 class=xl633834 width = 247 style='border-left:none;width:185pt'>Country</td>
 </tr>
 <tr height = 20 style='height:15.0pt'>
  <td height = 20 class=xl153834 style = 'height:15.0pt' ></ td >
  
    < td colspan=3 rowspan=2 class=xl643834 width = 296 style='width:222pt'>%%CompanyName%%</td>
  <td colspan = 3 rowspan=2 class=xl643834 width = 352 style='width:264pt'>%%ContactPerson%%</td>
  <td colspan = 3 rowspan=2 class=xl643834 width = 247 style='width:185pt'>%%Email%%</td>
  <td colspan = 3 rowspan= 2 class=xl643834 width = 452 style='width:339pt'>%%Cellphone%%</td>
  <td colspan = 3 rowspan=2 class=xl643834 width = 247 style='width:185pt'>%%Country%%</td>
 </tr>
 <tr height = 20 style='height:15.0pt'>
  <td height = 20 class=xl153834 style = 'height:15.0pt' ></ td >
  
   </ tr >
  
   < tr height=20 style='height:15.0pt'>
  <td height = 20 class=xl153834 style = 'height:15.0pt' ></ td >
  
    < td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
 </tr>
 <tr height = 21 style='height:15.75pt'>
  <td height = 21 class=xl153834 style = 'height:15.75pt' ></ td >
  
    < td colspan=5 class=xl633834 width = 479 style='width:359pt'>Cargo Information</td>
  <td colspan = 5 class=xl633834 width = 535 style='border-left:none;width:401pt'>Cargo
      Reception Information</td>
  <td colspan = 5 class=xl633834 width = 580 style='border-left:none;width:435pt'>Cargo
      Delivery Information</td>
 </tr>
 <tr height = 21 style='height:15.75pt'>
  <td height = 21 class=xl153834 style = 'height:15.75pt' ></ td >
  
    < td class=xl653834 style = 'border-top:none' > Total equipment:</td>
  <td colspan = 4 class=xl643834 width = 310 style='border-left:none;width:232pt'>%%TotalEquipment%%</td>
  <td class=xl653834 style = 'border-top:none;border-left:none' > Country of
      Origin:</td>
  <td colspan = 4 class=xl643834 width = 366 style='border-left:none;width:274pt'>%%CountryOfOrigin%%</td>
  <td class=xl653834 style = 'border-top:none;border-left:none' > Country of
      Destination:</td>
  <td colspan = 4 class=xl643834 width = 366 style='border-left:none;width:274pt'>%%CountryOfDestination%%</td>
 </tr>
 <tr height = 21 style= 'height:15.75pt' >
    
      < td height= 21 class=xl153834 style = 'height:15.75pt' ></ td >
      
        < td class=xl653834 style = 'border-top:none' > Weight(Kg):</td>
  <td colspan = 4 class=xl643834 width = 310 style='border-left:none;width:232pt'>%%Weight%%</td>
  <td class=xl653834 style = 'border-top:none;border-left:none' > Region of Origin:</td>
  <td colspan = 4 class=xl643834 width = 366 style='border-left:none;width:274pt'>%%RegionOfOrigin%%</td>
  <td class=xl653834 style = 'border-top:none;border-left:none' > Region of
      Destination:</td>
  <td colspan = 4 class=xl643834 width = 366 style='border-left:none;width:274pt'>%%RegionOfDestination%%</td>
 </tr>
 <tr height = 21 style= 'height:15.75pt' >
    
      < td height= 21 class=xl153834 style = 'height:15.75pt' ></ td >
      
        < td class=xl653834 style = 'border-top:none' > Incoterms:</td>
  <td colspan = 4 class=xl643834 width = 310 style='border-left:none;width:232pt'>%%Incoterms%%</td>
  <td class=xl653834 style = 'border-top:none;border-left:none' > City:</td>
  <td colspan = 4 class=xl683834 width = 366 style='border-left:none;width:274pt'>%%CityOfOrigin%%</td>
  <td class=xl653834 style = 'border-top:none;border-left:none' > City:</td>
  <td colspan = 4 class=xl683834 width = 366 style='border-left:none;width:274pt'>%%CityOfDestination%%</td>
 </tr>
 <tr height = 21 style='height:15.75pt'>
  <td height = 21 class=xl153834 style = 'height:15.75pt' ></ td >
  
    < td class=xl653834 style = 'border-top:none' > Type of commodity:</td>
  <td colspan = 4 class=xl643834 width = 310 style='border-left:none;width:232pt'>%%TypeOfCommodity%%</td>
  <td class=xl653834 style = 'border-top:none;border-left:none' > Receipt:</td>
  <td colspan = 4 class=xl643834 width = 366 style='border-left:none;width:274pt'>%%Receipt%%</td>
  <td class=xl653834 style = 'border-top:none;border-left:none' > Delivery:</td>
  <td colspan = 4 class=xl643834 width = 366 style='border-left:none;width:274pt'>Delivery</td>
 </tr>
 <tr height = 21 style='height:15.75pt'>
  <td height = 21 class=xl153834 style = 'height:15.75pt' ></ td >
  
    < td class=xl653834 style = 'border-top:none' > Cubic feets:</td>
  <td colspan = 4 class=xl643834 width = 310 style='border-left:none;width:232pt'>%%Cubicfeets%%</td>
  <td class=xl653834 style = 'border-top:none;border-left:none' > Departure Date:</td>
  <td colspan = 4 class=xl643834 width = 366 style='border-left:none;width:274pt'>%%DepartureDate%%</td>
  <td class=xl653834 style = 'border-top:none;border-left:none' > Arroval Date:</td>
  <td colspan = 4 class=xl643834 width = 366 style='border-left:none;width:274pt'>%%ArrivalDate%%</td>
 </tr>
 <tr height = 20 style='height:15.0pt'>
  <td height = 20 class=xl153834 style = 'height:15.0pt' ></ td >
  
    < td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
 </tr>
 <tr height = 21 style='height:15.75pt'>
  <td height = 21 class=xl153834 style = 'height:15.75pt' ></ td >
  
    < td colspan=15 class=xl633834 width = 1594 style='width:1195pt'>OCEAN CARRIERS
  / CARGO AGENTS / FREIGHT FORWARDERS</td>
 </tr>
 <tr height = 20 style= 'height:15.0pt' >
  < td height= 20 class=xl153834 style = 'height:15.0pt' ></ td >
  
    < td colspan=2 rowspan=3 class=xl633834 width = 233 style='width:175pt'>Services</td>
  <td rowspan = 2 class=xl633834 width = 63 style='border-top:none;width:47pt'>Unit</td>
  <td colspan = 2 rowspan=2 class=xl633834 width = 183 style='width:137pt'>Provider____________</td>
  <td colspan = 2 rowspan=2 class=xl633834 width = 233 style='width:175pt'>Provider____________</td>
  <td colspan = 2 rowspan=2 class=xl633834 width = 183 style='width:137pt'>Provider____________</td>
  <td colspan = 2 rowspan=2 class=xl633834 width = 333 style='width:250pt'>Provider____________</td>
  <td colspan = 2 rowspan=2 class=xl633834 width = 183 style='width:137pt'>Provider____________</td>
  <td colspan = 2 rowspan=2 class=xl633834 width = 183 style='width:137pt'>Provider____________</td>
 </tr>
 <tr height = 20 style='height:15.0pt'>
  <td height = 20 class=xl153834 style = 'height:15.0pt' ></ td >
  
   </ tr >
  
   < tr height=21 style='height:15.75pt'>
  <td height = 21 class=xl153834 style = 'height:15.75pt' ></ td >
  
    < td class=xl633834 width = 63 style='border-top:none;border-left:none;
  width:47pt'>&nbsp;</td>
  <td class=xl633834 width = 119 style='border-top:none;border-left:none;
  width:89pt'>Price/Unit</td>
  <td class=xl633834 width = 64 style='border-top:none;border-left:none;
  width:48pt'>Total</td>
  <td class=xl633834 width = 169 style='border-top:none;border-left:none;
  width:127pt'>Price/Unit</td>
  <td class=xl633834 width = 64 style='border-top:none;border-left:none;
  width:48pt'>Total</td>
  <td class=xl633834 width = 119 style='border-top:none;border-left:none;
  width:89pt'>Price/Unit</td>
  <td class=xl633834 width = 64 style='border-top:none;border-left:none;
  width:48pt'>Total</td>
  <td class=xl633834 width = 119 style='border-top:none;border-left:none;
  width:89pt'>Price/Unit</td>
  <td class=xl633834 width = 214 style='border-top:none;border-left:none;
  width:161pt'>Total</td>
  <td class=xl633834 width = 119 style='border-top:none;border-left:none;
  width:89pt'>Price/Unit</td>
  <td class=xl633834 width = 64 style='border-top:none;border-left:none;
  width:48pt'>Total</td>
  <td class=xl633834 width = 119 style='border-top:none;border-left:none;
  width:89pt'>Price/Unit</td>
  <td class=xl633834 width = 64 style='border-top:none;border-left:none;
  width:48pt'>Total</td>
 </tr>
 <tr height = 21 style='height:15.75pt'>
  <td height = 21 class=xl153834 style = 'height:15.75pt' ></ td >
  
    < td colspan=2 class=xl653834>Sea Freight</td>
  <td class=xl643834 width = 63 style='border-top:none;border-left:none;
  width:47pt'>&nbsp;</td>
  <td class=xl643834 width = 119 style='border-top:none;border-left:none;
  width:89pt'>&nbsp;</td>
  <td class=xl643834 width = 64 style='border-top:none;border-left:none;
  width:48pt'>&nbsp;</td>
  <td class=xl643834 width = 169 style='border-top:none;border-left:none;
  width:127pt'>&nbsp;</td>
  <td class=xl643834 width = 64 style='border-top:none;border-left:none;
  width:48pt'>&nbsp;</td>
  <td class=xl643834 width = 119 style='border-top:none;border-left:none;
  width:89pt'>&nbsp;</td>
  <td class=xl643834 width = 64 style='border-top:none;border-left:none;
  width:48pt'>&nbsp;</td>
  <td class=xl643834 width = 119 style='border-top:none;border-left:none;
  width:89pt'>&nbsp;</td>
  <td class=xl643834 width = 214 style='border-top:none;border-left:none;
  width:161pt'>&nbsp;</td>
  <td class=xl643834 width = 119 style='border-top:none;border-left:none;
  width:89pt'>&nbsp;</td>
  <td class=xl643834 width = 64 style='border-top:none;border-left:none;
  width:48pt'>&nbsp;</td>
  <td class=xl643834 width = 119 style='border-top:none;border-left:none;
  width:89pt'>&nbsp;</td>
  <td class=xl643834 width = 64 style='border-top:none;border-left:none;
  width:48pt'>&nbsp;</td>
 </tr>
 ";
            excel.OceanCarriersInfo.LogisticServices.ForEach(item => {
                excelString += @$"<tr height = 21 style='height:15.75pt'>
                                  < td height = 21 class=xl153834 style = 'height:15.75pt' ></ td >
                                  < td colspan=2 class=xl653834>{item.Name}</td>
                                  <td class=xl643834 width = 63 style='border-top:none;border-left:none;
                                  width:47pt'>{item.Unit}</td>
                                  <td class=xl643834 width = 119 style='border-top:none;border-left:none;
                                  width:89pt'>{item.PriceUnit}</td>
                                  <td class=xl643834 width = 64 style='border-top:none;border-left:none;
                                  width:48pt'>{item.Total}</td>
                                  <td class=xl643834 width = 169 style='border-top:none;border-left:none;
                                  width:127pt'>{item.PriceUnit};</td>
                                  <td class=xl643834 width = 64 style='border-top:none;border-left:none;
                                  width:48pt'>{item.Total}</td>
                                  <td class=xl643834 width = 119 style='border-top:none;border-left:none;
                                  width:89pt'>{item.PriceUnit}</td>
                                  <td class=xl643834 width = 64 style='border-top:none;border-left:none;
                                  width:48pt'>{item.Total}</td>
                                  <td class=xl643834 width = 119 style='border-top:none;border-left:none;
                                  width:89pt'>{item.PriceUnit}</td>
                                  <td class=xl643834 width = 214 style='border-top:none;border-left:none;
                                  width:161pt'>{item.Total}</td>
                                  <td class=xl643834 width = 119 style='border-top:none;border-left:none;
                                  width:89pt'>{item.PriceUnit}</td>
                                  <td class=xl643834 width = 64 style='border-top:none;border-left:none;
                                  width:48pt'>{item.Total}</td>
                                  <td class=xl643834 width = 119 style='border-top:none;border-left:none;
                                  width:89pt'>{item.PriceUnit}</td>
                                  <td class=xl643834 width = 64 style='border-top:none;border-left:none;
                                  width:48pt'>{item.Total}</td>
                                 </tr>";
            });
            excelString += @$"<tr height = 21 style='height:15.75pt'>
  <td height = 21 class=xl153834 style = 'height:15.75pt' ></ td >
  
    < td class=xl703834 style = 'border-top:none' > Subtotal </ td >
   
     < td class=xl653834 style = 'border-top:none;border-left:none' > {excel.OceanCarriersInfo.SubTotal}</td>
  <td class=xl643834 width = 63 style='border-top:none;border-left:none;
  width:47pt'>&nbsp;</td>
  <td class=xl643834 width = 119 style='border-top:none;border-left:none;
  width:89pt'>&nbsp;</td>
  <td class=xl643834 width = 64 style='border-top:none;border-left:none;
  width:48pt'>&nbsp;</td>
  <td class=xl643834 width = 169 style='border-top:none;border-left:none;
  width:127pt'>&nbsp;</td>
  <td class=xl643834 width = 64 style='border-top:none;border-left:none;
  width:48pt'>&nbsp;</td>
  <td class=xl643834 width = 119 style='border-top:none;border-left:none;
  width:89pt'>&nbsp;</td>
  <td class=xl643834 width = 64 style='border-top:none;border-left:none;
  width:48pt'>&nbsp;</td>
  <td class=xl643834 width = 119 style='border-top:none;border-left:none;
  width:89pt'>&nbsp;</td>
  <td class=xl643834 width = 214 style='border-top:none;border-left:none;
  width:161pt'>&nbsp;</td>
  <td class=xl643834 width = 119 style='border-top:none;border-left:none;
  width:89pt'>&nbsp;</td>
  <td class=xl643834 width = 64 style='border-top:none;border-left:none;
  width:48pt'>&nbsp;</td>
  <td class=xl643834 width = 119 style='border-top:none;border-left:none;
  width:89pt'>&nbsp;</td>
  <td class=xl643834 width = 64 style='border-top:none;border-left:none;
  width:48pt'>&nbsp;</td>
 </tr>
 <tr height = 21 style='height:15.75pt'>
  <td height = 21 class=xl153834 style = 'height:15.75pt' ></ td >
  
    < td class=xl703834 style = 'border-top:none' > Taxes </ td >
   
     < td class=xl653834 style = 'border-top:none;border-left:none' > {excel.OceanCarriersInfo.Taxes} </td>
  <td class=xl643834 width = 63 style='border-top:none;border-left:none;
  width:47pt'>&nbsp;</td>
  <td class=xl643834 width = 119 style='border-top:none;border-left:none;
  width:89pt'>&nbsp;</td>
  <td class=xl643834 width = 64 style='border-top:none;border-left:none;
  width:48pt'>&nbsp;</td>
  <td class=xl643834 width = 169 style='border-top:none;border-left:none;
  width:127pt'>&nbsp;</td>
  <td class=xl643834 width = 64 style='border-top:none;border-left:none;
  width:48pt'>&nbsp;</td>
  <td class=xl643834 width = 119 style='border-top:none;border-left:none;
  width:89pt'>&nbsp;</td>
  <td class=xl643834 width = 64 style='border-top:none;border-left:none;
  width:48pt'>&nbsp;</td>
  <td class=xl643834 width = 119 style='border-top:none;border-left:none;
  width:89pt'>&nbsp;</td>
  <td class=xl643834 width = 214 style='border-top:none;border-left:none;
  width:161pt'>&nbsp;</td>
  <td class=xl643834 width = 119 style='border-top:none;border-left:none;
  width:89pt'>&nbsp;</td>
  <td class=xl643834 width = 64 style='border-top:none;border-left:none;
  width:48pt'>&nbsp;</td>
  <td class=xl643834 width = 119 style='border-top:none;border-left:none;
  width:89pt'>&nbsp;</td>
  <td class=xl643834 width = 64 style='border-top:none;border-left:none;
  width:48pt'>&nbsp;</td>
 </tr>
 <tr height = 21 style='height:15.75pt'>
  <td height = 21 class=xl153834 style = 'height:15.75pt' ></ td >
  
    < td class=xl703834 style = 'border-top:none' > Total </ td >
   
     < td class=xl653834 style = 'border-top:none;border-left:none' > {excel.OceanCarriersInfo.Total} </td>
  <td class=xl643834 width = 63 style='border-top:none;border-left:none;
  width:47pt'>&nbsp;</td>
  <td class=xl643834 width = 119 style='border-top:none;border-left:none;
  width:89pt'>&nbsp;</td>
  <td class=xl643834 width = 64 style='border-top:none;border-left:none;
  width:48pt'>&nbsp;</td>
  <td class=xl643834 width = 169 style='border-top:none;border-left:none;
  width:127pt'>&nbsp;</td>
  <td class=xl643834 width = 64 style='border-top:none;border-left:none;
  width:48pt'>&nbsp;</td>
  <td class=xl643834 width = 119 style='border-top:none;border-left:none;
  width:89pt'>&nbsp;</td>
  <td class=xl643834 width = 64 style='border-top:none;border-left:none;
  width:48pt'>&nbsp;</td>
  <td class=xl643834 width = 119 style='border-top:none;border-left:none;
  width:89pt'>&nbsp;</td>
  <td class=xl643834 width = 214 style='border-top:none;border-left:none;
  width:161pt'>&nbsp;</td>
  <td class=xl643834 width = 119 style='border-top:none;border-left:none;
  width:89pt'>&nbsp;</td>
  <td class=xl643834 width = 64 style='border-top:none;border-left:none;
  width:48pt'>&nbsp;</td>
  <td class=xl643834 width = 119 style='border-top:none;border-left:none;
  width:89pt'>&nbsp;</td>
  <td class=xl643834 width = 64 style='border-top:none;border-left:none;
  width:48pt'>&nbsp;</td>
 </tr>
 <tr height = 20 style='height:15.0pt'>
  <td height = 20 class=xl153834 style = 'height:15.0pt' ></ td >
  
    < td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
 </tr>
 <tr height = 21 style='height:15.75pt'>
  <td height = 21 class=xl153834 style = 'height:15.75pt' ></ td >
  
    < td colspan=15 class=xl633834 width = 1594 style='width:1195pt'>Special
       Instructions</td>
 </tr>
 <tr height = 20 style= 'height:15.0pt' >
     
       < td height= 20 class=xl153834 style = 'height:15.0pt' ></ td >
       
         < td colspan=15 rowspan=3 class=xl643834 width = 1594 style='width:1195pt'>%%SpecialInstructions%%</td>
 </tr>
 <tr height = 20 style='height:15.0pt'>
  <td height = 20 class=xl153834 style = 'height:15.0pt' ></ td >
  
   </ tr >
  
   < tr height=20 style='height:15.0pt'>
  <td height = 20 class=xl153834 style = 'height:15.0pt' ></ td >
  
   </ tr >
  
   < tr height=20 style='height:15.0pt'>
  <td height = 20 class=xl153834 style = 'height:15.0pt' ></ td >
  
    < td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
 </tr>
 <tr height = 20 style='height:15.0pt'>
  <td height = 20 class=xl153834 style = 'height:15.0pt' ></ td >
  
    < td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
  <td class=xl153834></td>
 </tr>
 <tr height = 21 style='height:15.75pt'>
  <td height = 21 class=xl153834 style = 'height:15.75pt' ></ td >
  
    < td colspan=15 class=xl643834 width = 1594 style='width:1195pt'>Inpectores de
  Carga: Qualitas Bureau; Contacto: Raul De Saint Malo; &nbsp;457 B Chame
  Street, Ancon; &nbsp;Panama, Rep of Panama; &nbsp;Telef: +507-203-8239;
  email: surveys @qualitasbureau.com ; www.qualitasbureau.com</td>
 </tr>
 <![if supportMisalignedColumns]>
 <tr height = 0 style='display:none'>
  <td width = 64 style='width:48pt'></td>
  <td width = 169 style='width:127pt'></td>
  <td width = 64 style='width:48pt'></td>
  <td width = 63 style='width:47pt'></td>
  <td width = 119 style='width:89pt'></td>
  <td width = 64 style='width:48pt'></td>
  <td width = 169 style='width:127pt'></td>
  <td width = 64 style='width:48pt'></td>
  <td width = 119 style='width:89pt'></td>
  <td width = 64 style='width:48pt'></td>
  <td width = 119 style='width:89pt'></td>
  <td width = 214 style='width:161pt'></td>
  <td width = 119 style='width:89pt'></td>
  <td width = 64 style='width:48pt'></td>
  <td width = 119 style='width:89pt'></td>
  <td width = 64 style='width:48pt'></td>
 </tr>
 <![endif]>
</table>

</div>


<!----------------------------->
<!--FINAL DE LOS RESULTADOS DEL ASISTENTE PARA PUBLICAR COMO PÁGINA WEB DE
EXCEL-->
<!----------------------------->
</body>

</html>
";
            return excelString;
        }
    }
}
