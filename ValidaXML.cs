using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.IO;
using System.Windows.Forms;
using System.Net;
using System.Xml.Schema;

namespace EmissorCTe.CS
{
    class ValidaXML
    {
        private int resultado;
        private string msgResultado;

        public int getResultado
        {
            get { return resultado; }
        }

        public string getMsgResultado
        {
            get { return msgResultado; }
        }

        public int validar(string pArquivo, string pSchema)
        /*
         *     Entradas:
         *         _arquivo : Nome do arquivo informado
         *         _schema  : Shema do arquivo XML em questão
         *         
         *     Retornos:
         *         Validar : 0 - Assinatura realizada com sucesso
         *                   1 - XML mal formado + "Linha: {0} Coluna:{1} Erro:{2}"
         */
        {
            resultado = 0;
            msgResultado = "Validação realizada com sucesso";

            // Create a new validating reader
            Stream s = new MemoryStream(ASCIIEncoding.Default.GetBytes(pArquivo));
            XmlValidatingReader oReader = new XmlValidatingReader(new XmlTextReader(new StreamReader(s)));

            // Create a schema collection, add the xsd to it
            XmlSchemaCollection oSchemaCollection = new XmlSchemaCollection();
            oSchemaCollection.Add("http://www.portalfiscal.inf.br/cte", pSchema);

            // Add the schema collection to the XmlValidatingReader
            oReader.Schemas.Add(oSchemaCollection);

            // Wire up the call back.  The ValidationEvent is fired when the
            // XmlValidatingReader hits an issue validating a section of the xml
            oReader.ValidationEventHandler += new ValidationEventHandler(getReaderValidationEventHandler);

            // Iterate through the xml document
            while (oReader.Read()) { if (resultado > 0) { break; } }

            return resultado;
        }

        private void getReaderValidationEventHandler(object pSender, ValidationEventArgs pException)
        {
            // Report back error information to the console...
            resultado = 1;
            msgResultado = "Linha: " + Convert.ToString(pException.Exception.LinePosition) +
                           " Coluna: " + Convert.ToString(pException.Exception.LineNumber) +
                           " Erro: " + pException.Exception.Message;
        }
    }
}
