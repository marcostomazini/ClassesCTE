using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Security.Cryptography.X509Certificates;
using System.Xml;
using System.Security.Cryptography.Xml;

namespace EmissorCTe.CS
{
    class AssinaXML
    {
        private int resultado;
        private string msgResultado;
        private XmlDocument xmlDoc;

        public string getResultado
        {
            get { return msgResultado; }
        }

        public string getMsgResultado
        {
            get { return msgResultado; }
        }

        public XmlDocument getXMLDocAssinado
        {
            get { return xmlDoc; }
        }

        public string getXMLStringAssinado
        {
            get { return xmlDoc.OuterXml; }
        }

        public int assinar(XmlDocument oDoc, string pRefUri, X509Certificate2 pCertificado)
        /*
         *     Entradas:
         *         pXMLString   : string XML a ser assinada
         *         pUri         : Referência da URI a ser assinada (Ex. infNFe
         *         pCertificado : certificado digital a ser utilizado na assinatura digital
         * 
         *     Retornos:
         *         Assinar : 0 - Assinatura realizada com sucesso
         *                   1 - Erro: Problema ao acessar o certificado digital - %exceção%
         *                   2 - Problemas no certificado digital
         *                   3 - XML mal formado + exceção
         *                   4 - A tag de assinatura %Uri% inexiste
         *                   5 - A tag de assinatura %Uri% não é unica
         *                   6 - Erro Ao assinar o documento - ID deve ser string %Uri(Atributo)%
         *                   7 - Erro: Ao assinar o documento - %exceção%
         * 
         *         XMLStringAssinado : string XML assinada
         * 
         *         XMLDocAssinado    : XMLDocument do XML assinado
         */
        {
            resultado = 0;
            msgResultado = "Assinatura realizada com sucesso";
            try
            {
                string pNome = "";
                if (pCertificado != null)
                {
                    pNome = pCertificado.Subject.ToString();
                }

                X509Certificate2 oX509Certificado = new X509Certificate2();
                X509Store oStore = new X509Store("MY", StoreLocation.CurrentUser);
                oStore.Open(OpenFlags.ReadOnly | OpenFlags.OpenExistingOnly);
                X509Certificate2Collection oCollection = (X509Certificate2Collection)oStore.Certificates;
                X509Certificate2Collection oCollection1 = (X509Certificate2Collection)oCollection.Find(X509FindType.FindBySubjectDistinguishedName, pNome, false);
                if (oCollection1.Count == 0)
                {
                    resultado = 2;
                    msgResultado = "Problemas no certificado digital";
                }
                else
                {
                    // certificado ok
                    oX509Certificado = oCollection1[0];
                    string x;
                    x = oX509Certificado.GetKeyAlgorithm().ToString();
                    try
                    {
                        // Verifica se a tag a ser assinada existe é única
                        int vQtdeRefUri = oDoc.GetElementsByTagName(pRefUri).Count;

                        if (vQtdeRefUri == 0)
                        {
                            //  a URI indicada não existe
                            resultado = 4;
                            msgResultado = "A tag de assinatura " + pRefUri.Trim() + " inexiste";
                        }
                        else
                        {
                            if (vQtdeRefUri > 1)
                            {
                                // existe mais de uma URI indicada
                                resultado = 5;
                                msgResultado = "A tag de assinatura " + pRefUri.Trim() + " não é unica";

                            }
                            else
                            {
                                try
                                {
                                    // Create a SignedXml object.
                                    SignedXml oSignedXML = new SignedXml(oDoc);

                                    // Add the key to the SignedXml document 
                                    oSignedXML.SigningKey = oX509Certificado.PrivateKey;

                                    // Create a reference to be signed
                                    Reference oReference = new Reference();

                                    // pega o uri que deve ser assinada
                                    XmlAttributeCollection oUri = oDoc.GetElementsByTagName(pRefUri).Item(0).Attributes;
                                    foreach (XmlAttribute oAtributo in oUri)
                                    {
                                        if (oAtributo.Name == "Id")
                                        {
                                            oReference.Uri = "#" + oAtributo.InnerText;
                                        }
                                    }

                                    // Add an enveloped transformation to the reference.
                                    XmlDsigEnvelopedSignatureTransform oEnv = new XmlDsigEnvelopedSignatureTransform();
                                    oReference.AddTransform(oEnv);

                                    XmlDsigC14NTransform oC14 = new XmlDsigC14NTransform();
                                    oReference.AddTransform(oC14);

                                    // Add the reference to the SignedXml object.
                                    oSignedXML.AddReference(oReference);

                                    // Create a new KeyInfo object
                                    KeyInfo oKeyInfo = new KeyInfo();

                                    // Load the certificate into a KeyInfoX509Data object
                                    // and add it to the KeyInfo object.
                                    oKeyInfo.AddClause(new KeyInfoX509Data(oX509Certificado));

                                    // Add the KeyInfo object to the SignedXml object.
                                    oSignedXML.KeyInfo = oKeyInfo;

                                    oSignedXML.ComputeSignature();

                                    // Get the XML representation of the signature and save
                                    // it to an XmlElement object.
                                    XmlElement oXMLDigitalSignature = oSignedXML.GetXml();

                                    // Append the element to the XML document.
                                    oDoc.DocumentElement.AppendChild(oDoc.ImportNode(oXMLDigitalSignature, true));
                                    xmlDoc = new XmlDocument();
                                    xmlDoc.PreserveWhitespace = false;
                                    xmlDoc = oDoc;
                                }
                                catch (Exception caught)
                                {
                                    resultado = 7;
                                    msgResultado = "Erro: Ao assinar o documento - " + caught.Message;
                                }
                            }
                        }
                    }
                    catch (Exception caught)
                    {
                        resultado = 3;
                        msgResultado = "Erro: XML mal formado - " + caught.Message;
                    }
                }
            }
            catch (Exception caught)
            {
                resultado = 1;
                msgResultado = "Erro: Problema ao acessar o certificado digital" + caught.Message;
            }

            return resultado;
        }
    }
}
