using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Security.Cryptography.X509Certificates;
using System.Xml;

namespace EmissorCTe.CS
{
    class Funcoes
    {
        public static string DigitoModulo11(string chave)
        {
            int digitoRetorno;
            int soma = 0;
            int resto = 0;
            int[] peso = { 4, 3, 2, 9, 8, 7, 6, 5 };

            for (int i = 0; i < chave.Length; i++)
            {
                soma += peso[i % 8] * (int.Parse(chave.Substring(i, 1)));
            }

            resto = soma % 11;
            if (resto == 0 || resto == 1)
            {
                digitoRetorno = 0;
            }
            else
            {
                digitoRetorno = 11 - resto;
            }
            return chave + digitoRetorno.ToString();
        }

        public static X509Certificate2 selecionaCertificado(X509Certificate2 pCertificado)
        {
            // Seleciona certificado do repositório MY do windows // marcos@tomazini.org
            Certificado vCertificado = new Certificado();
            pCertificado = vCertificado.BuscaNome("");

            if (vCertificado.getResultado != 0)
            {
                // Mensagem
                MessageBox.Show(vCertificado.getMsgResultado, "Atenção",
                   MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }
            else
            {
                return pCertificado;
            }
        }

        public static bool assinaXML(AssinaXML pAssina, XmlDocument XML, X509Certificate2 certificado)
        {
            // Realiza assinatura
            if (pAssina.assinar(XML, "infCte", certificado) == 0)
            {
                Constantes.XML_Assinado = pAssina.getXMLStringAssinado;                
                return true;
            }
            else if (pAssina.assinar(XML, "infCanc", certificado) == 0)
            {
                Constantes.XML_Cancelamento_Assinado = pAssina.getXMLStringAssinado;
                return true;
            }
            else if (pAssina.assinar(XML, "infInut", certificado) == 0)
            {
                Constantes.XML_Inutilizacao_Assinado = pAssina.getXMLStringAssinado;
                return true;
            }
            else
            {
                MessageBox.Show(pAssina.getMsgResultado, "Atenção",
                   MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
        }

        private string getXMLValorTagRec(string pTagPai, string pTag, XmlNodeList pChild)
        {
            string vConteudo = "";
            //Percorre todos os elementos filhos que existem no elemento raiz 
            for (int i = 0; i < pChild.Count; i++)
            {
                if (pTagPai == "")
                {
                    if (pChild.Item(i).Name == pTag)
                    {
                        vConteudo = pChild.Item(i).InnerText;
                        break;
                    }
                }
                else
                {
                    if (pChild.Item(i).Name == pTagPai)
                    {
                        //Percorre todos os elementos que estiverem dentro do elemento filho atual  // marcos@tomazini.org
                        for (int a = 0; a < pChild.Item(i).ChildNodes.Count; a++)
                        {
                            if (pChild.Item(i).ChildNodes.Item(a).Name == pTag)
                            {
                                vConteudo = pChild.Item(i).ChildNodes.Item(a).InnerText;
                                break;
                            }
                        }
                    }
                    else
                    {
                        vConteudo = getXMLValorTagRec(pTagPai, pTag, pChild.Item(i).ChildNodes);
                        if (vConteudo != "") { break; }
                    }
                }
            }
            return vConteudo;
        }

        public string getXMLValorTag(string pTagPai, string pTag, string pXMLString)
        {
            string vConteudo = "";
            if (pXMLString != null)
            {
                XmlDocument oDoc = new XmlDocument();
                oDoc.PreserveWhitespace = false;
                // Carrega XML
                oDoc.LoadXml(pXMLString);
                // Cria uma instância XmlElement na qual atribuímos a raiz do documento 
                XmlElement oElem = oDoc.DocumentElement;
                // Pega valor da Tag
                vConteudo = getXMLValorTagRec(pTagPai, pTag, oElem.ChildNodes);
            }
            return vConteudo;
        }

    }
}
