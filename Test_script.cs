using System;
using System.Collections.Generic;

//string strExportPath = FCTools.ScriptContext.Project.EnvironmentVariables.Get("ExportPath");  //"C:\\ABBYY\\Export\\"; // два слэша на конце обязательны


string connStr = FCTools.ScriptContext.Project.EnvironmentVariables.Get("XMLExchange"); //"Provider = SQLOLEDB; Data Source=10.11.1.12;Initial Catalog=ExchangeDB;Persist Security Info=True;User ID=xmlexchange;Password=Aa1234";



    ABBYY.FlexiCapture.IExportImageSavingOptions imgExp = FCTools.NewImageSavingOptions();
    imgExp.ColorType = FCTools.ScriptContext.Project.EnvironmentVariables.Get("ColorType");//"GrayScale";
    imgExp.Format = FCTools.ScriptContext.Project.EnvironmentVariables.Get("Format");//"pdf-a-s";
    imgExp.Resolution = Convert.ToInt32(FCTools.ScriptContext.Project.EnvironmentVariables.Get("Resolution"));//300;
    imgExp.Quality = Convert.ToInt32(FCTools.ScriptContext.Project.EnvironmentVariables.Get("Quality"));
    //imgExp.ImageCompressionType = ICT_Auto; //Convert.ToInt32(FCTools.ScriptContext.Project.EnvironmentVariables.Get("ImageCompressionType"));
    imgExp.UseMRC = true;
    imgExp.ShouldOverwrite = true;


//загрузка маппинга





//string connStr2 = FCTools.ScriptContext.Project.EnvironmentVariables.Get("XMLExchange");
//Dictionary<string, string> mapping = Ensol.Flexi.Common.Data.GetDocumentMappingServer(connStr2, "");    

//Поиск главного документа
/*string xmlResultParentDoc = "";
string ParentDocName = "";

for (int i = 0; i < Document.Sections.Count ; i++)
{
     if (Document.Sections[i].Name == "Счет-фактура") 
     {
        xmlResultParentDoc = Document.Sections[i].Field("Barcode").Text;
        ParentDocName = Document.Sections[i].Name;
        break;   
     }
     else if (Document.Sections[i].Name == "УПД" || Document.Sections[i].Name == "ТОРГ12" || Document.Sections[i].Name == "Акт")
    {
        xmlResultParentDoc = Document.Sections[i].Field("Barcode").Text;
        ParentDocName = Document.Sections[i].Name;
    }   
}*/

//xmlResultParentDoc = "<ListItem  FieldName=\"" + "ParentDoc" + "\">" + xmlResultParentDoc + "</ListItem>";


//заполнение данных из OEBS
string xmlResultZP = ""; 
for (int i = 0; i < Document.Sections.Count ; i++)
{
    xmlResultZP = Ensol.Flexi.Yandex.Data.GetOEBSID("http://localhost:8080/flexi/Ensol.Flexi.Service.svc/http/", Document.Sections[i].Field("Barcode").Text);
}


// по секциям
for (int i = 0; i < Document.Sections.Count ; i++)
{
    string xmlResult = "";
    // генерим XML- файл

        for (int j = 0; j < Document.Sections[i].Children.Count ; j++) // по полям
        {
            if ( Document.Sections[i].Children[j].Text.Length != 0 )
            {
                xmlResult += "<ListItem  FieldName=\"" + Document.Sections[i].Children[j].Name + "\">" + Document.Sections[i].Children[j].Text + "</ListItem>";
            }
        }
        xmlResultZP += xmlResult;  
        xmlResult += "</Element>";
       
    for (int j = 0; j < Document.Pages.Count; j++) // по всем страницам - делаем неэкспортируемыми
    {
        Document.Pages[j].ExcludedFromDocumentImage = true;
    }
    for (int j = 0; j < Document.Sections[i].Regions.Count; j++) // по всем регионам данной секции - смотрим страницу и помечаем как экспортируемую
    {
        Document.Pages[Document.Sections[i].Regions[j].Page.Index].ExcludedFromDocumentImage = false;
    }    
    // перебираем все страницы после секции, и если шаблон секции пустой - то это приложение и также помечаем как экспортируемую
    bool isEnd = false;
    int k = Document.Sections[i].Regions[Document.Sections[i].Regions.Count -1].Page.Index+1;
    while ( k < Document.Pages.Count && isEnd == false)
    {
        if ( Document.Pages[k].SectionName == "" )
        {
            Document.Pages[k].ExcludedFromDocumentImage = false;
        }
        else
        {
            isEnd = true;
        }
        k++;
    } 
    //}


    byte[] fileByteArray = Document.SaveAsStream(imgExp);
    //Document.SaveAs(strExportPath + Document.Sections[i].Field("Barcode").Text + FCTools.ScriptContext.Project.EnvironmentVariables.Get("Extension"), imgExp);
    // Загрузили pdf-файл
    //byte[] fileByteArray = System.IO.File.ReadAllBytes(strExportPath + Document.Sections[i].Field("Barcode").Text + FCTools.ScriptContext.Project.EnvironmentVariables.Get("Extension"));
    
    
    // Экспорт в XMLExchange
    using (System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection(connStr))
    {
        System.Data.OleDb.OleDbCommand dbCommand = conn.CreateCommand();
        dbCommand.CommandText = "INSERT INTO XmlExchangeUploadScan (Barcode, ExportDate, XMLDocument, IMPORTED, BinAttachment, BinAttachmentName,DestListName) VALUES (?, GETDATE(), ?, 0, ?, ?, ?)";
        
        dbCommand.Connection = conn;
        conn.Open();

        dbCommand.Parameters.Add("Barcode", System.Data.OleDb.OleDbType.VarChar, 100).Value = Document.Sections[i].Field("Barcode").Text;
        dbCommand.Parameters.Add("XMLDocument", System.Data.OleDb.OleDbType.VarChar).Value = xmlResult;
        //fileByteArray;
        dbCommand.Parameters.Add("BinAttachment", System.Data.OleDb.OleDbType.VarBinary).Value = fileByteArray;
        dbCommand.Parameters.Add("BinAttachmentName", System.Data.OleDb.OleDbType.VarChar, 30).Value = Document.Sections[i].Field("Barcode").Text + FCTools.ScriptContext.Project.EnvironmentVariables.Get("Extension");
        dbCommand.Parameters.Add("DestListName", System.Data.OleDb.OleDbType.VarChar).Value = DestListName;
        dbCommand.ExecuteNonQuery();
    }
        
}








/*<?xml version="1.0" encoding="windows-1251" ?> 
<Element BarCode="<Штриховой код документа>” ContentType="<Тип документа SP>"> 
<ListItem  fieldName="<Наименование поля в соответствии с меппингом>"> [Значение] </ListItem>
………
</Element>*/




