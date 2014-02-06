using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Windows.Forms;
using System.Data.SqlServerCe;
using EmissorCTe.DS;
using System.Data;
using System.IO;

namespace EmissorCTe.CS
{
    class GeraXML
    {
        private string RemoveCaracterEspeciais(string texto)
        {
            string textor = "";

            for (int i = 0; i < texto.Length; i++)
            {
                if (texto[i].ToString() == "ã") textor += "a";
                else if (texto[i].ToString() == "á") textor += "a";
                else if (texto[i].ToString() == "à") textor += "a";
                else if (texto[i].ToString() == "â") textor += "a";
                else if (texto[i].ToString() == "ä") textor += "a";
                else if (texto[i].ToString() == "é") textor += "e";
                else if (texto[i].ToString() == "è") textor += "e";
                else if (texto[i].ToString() == "ê") textor += "e";
                else if (texto[i].ToString() == "ë") textor += "e";
                else if (texto[i].ToString() == "í") textor += "i";
                else if (texto[i].ToString() == "ì") textor += "i";
                else if (texto[i].ToString() == "ï") textor += "i";
                else if (texto[i].ToString() == "õ") textor += "o";
                else if (texto[i].ToString() == "ó") textor += "o";
                else if (texto[i].ToString() == "ò") textor += "o";
                else if (texto[i].ToString() == "ö") textor += "o";
                else if (texto[i].ToString() == "ú") textor += "u";
                else if (texto[i].ToString() == "ù") textor += "u";
                else if (texto[i].ToString() == "ü") textor += "u";
                else if (texto[i].ToString() == "ç") textor += "c";
                else if (texto[i].ToString() == "Ã") textor += "A";
                else if (texto[i].ToString() == "Á") textor += "A";
                else if (texto[i].ToString() == "À") textor += "A";
                else if (texto[i].ToString() == "Â") textor += "A";
                else if (texto[i].ToString() == "Ä") textor += "A";
                else if (texto[i].ToString() == "É") textor += "E";
                else if (texto[i].ToString() == "È") textor += "E";
                else if (texto[i].ToString() == "Ê") textor += "E";
                else if (texto[i].ToString() == "Ë") textor += "E";
                else if (texto[i].ToString() == "Í") textor += "I";
                else if (texto[i].ToString() == "Ì") textor += "I";
                else if (texto[i].ToString() == "Ï") textor += "I";
                else if (texto[i].ToString() == "Õ") textor += "O";
                else if (texto[i].ToString() == "Ó") textor += "O";
                else if (texto[i].ToString() == "Ò") textor += "O";
                else if (texto[i].ToString() == "Ö") textor += "O";
                else if (texto[i].ToString() == "Ú") textor += "U";
                else if (texto[i].ToString() == "Ù") textor += "U";
                else if (texto[i].ToString() == "Ü") textor += "U";
                else if (texto[i].ToString() == "Ç") textor += "C";
                // especiais
                /*else if (texto[i].ToString() == "<") textor += "&lt;";
                else if (texto[i].ToString() == ">") textor += "&gt;";
                else if (texto[i].ToString() == "&") textor += "&amp;";
                else if (texto[i].ToString() == "\"") textor += "&quot;";
                else if (texto[i].ToString() == "'") textor += "&#39;";*/
                else textor += texto[i];
            }
            return textor;
        }

        public void GeraInutilizacaoXml(documento Dados, string chave_acesso, string Motivo)
        {
            try
            {
                string Ano = String.Format("{0:yy}", Dados.XML.Rows[0]["dhEmi"]);
                string XML = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><inutCTe xmlns=\"http://www.portalfiscal.inf.br/cte\" versao=\"1.04\"><infInut Id=\"ID" + chave_acesso + "\"><tpAmb>" +
                    Dados.XML.Rows[0]["tpAmb"].ToString() + "</tpAmb><xServ>INUTILIZAR</xServ><cUF>" +
                    Constantes.cUF + "</cUF><ano>" + Ano +
                    "</ano><CNPJ>" + Dados.XML.Rows[0]["emit_CNPJ"].ToString().Replace(".", "").Replace("-", "").Replace("/", "") + "</CNPJ><mod>" +
                    Dados.XML.Rows[0]["mod"].ToString() + "</mod><serie>" + Dados.XML.Rows[0]["serie"].ToString() + "</serie><nCTIni>" +
                    Convert.ToInt32(chave_acesso.ToString().Substring(21, 9)) + "</nCTIni><nCTFin>" +
                    Convert.ToInt32(chave_acesso.ToString().Substring(30, 9)) + "</nCTFin><xJust>" +
                    Motivo + "</xJust></infInut></inutCTe>";

                XmlDocument xdoc = new XmlDocument();
                xdoc.LoadXml(XML);

                Constantes.XML_Inutilizacao = xdoc;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao Gerar XML de Inutilizacao\n" + ex.ToString());
            }
        }

        public void GerarCancelamentoXml(documento Dados, string chave_acesso, string Motivo)
        {
            try
            {
                string XML = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><cancCTe xmlns=\"http://www.portalfiscal.inf.br/cte\" versao=\"1.04\"><infCanc Id=\"ID" + Dados.Lote.Rows[0]["chCTe"].ToString() + "\"><tpAmb>" + Dados.Lote.Rows[0]["tpAmb"].ToString() + "</tpAmb><xServ>CANCELAR</xServ><chCTe>" +
                   Dados.Lote.Rows[0]["chCTe"].ToString() + "</chCTe><nProt>" + Dados.Lote.Rows[0]["nProt"].ToString() + "</nProt><xJust>" + Motivo + "</xJust></infCanc></cancCTe>";
                XmlDocument xdoc = new XmlDocument();
                xdoc.LoadXml(XML);

                Constantes.XML_Cancelamento = xdoc;

            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao Gerar XML de Cancelamento\n" + ex.ToString());
            }
        }

        public void GerarArquivoXml(documento Dados, int id_documento, string chave_acesso)
        {
            try
            {
                StringWriter sw = new StringWriter();
                XmlTextWriter _xml = new XmlTextWriter(sw);

                _xml.WriteStartDocument();

                #region CTe
                _xml.WriteStartElement("CTe", "http://www.portalfiscal.inf.br/cte");
                #region infCte
                _xml.WriteStartElement("infCte"); //Grupo que contém as informações da CT-e;

                _xml.WriteStartAttribute("Id");
                _xml.WriteString("CTe" + chave_acesso);
                _xml.WriteEndAttribute();

                _xml.WriteStartAttribute("versao");////Versão do leiaute (v2.0),Atributos do Nó
                _xml.WriteString("1.04");//vesao da nfe.
                _xml.WriteEndAttribute();

                #region ide
                _xml.WriteStartElement("ide");

                _xml.WriteStartElement("cUF");
                _xml.WriteString(Dados.XML.Rows[0]["cUF"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("cCT");
                _xml.WriteString(Dados.XML.Rows[0]["cCT"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("CFOP");
                _xml.WriteString(Dados.XML.Rows[0]["CFOP"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("natOp");
                _xml.WriteString(Dados.XML.Rows[0]["natOp"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("forPag");
                _xml.WriteString(Dados.XML.Rows[0]["forPag"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("mod");
                _xml.WriteString(Dados.XML.Rows[0]["mod"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("serie");
                _xml.WriteString(Dados.XML.Rows[0]["serie"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("nCT");
                _xml.WriteString(Dados.XML.Rows[0]["nCT"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("dhEmi");
                // 2011-08-20T18:24:19 - "yyyy-MM-ddThh:mm:ss"                
                _xml.WriteString(String.Format("{0:s}", Dados.XML.Rows[0]["dhEmi"]));
                _xml.WriteEndElement();

                _xml.WriteStartElement("tpImp");
                _xml.WriteString(Dados.XML.Rows[0]["tpImp"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("tpEmis");
                _xml.WriteString(Dados.XML.Rows[0]["tpEmis"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("cDV");
                _xml.WriteString(chave_acesso.Substring(43, 1)); //ultimo n* da chave = digito
                _xml.WriteEndElement();

                _xml.WriteStartElement("tpAmb");
                _xml.WriteString(Dados.XML.Rows[0]["tpAmb"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("tpCTe");
                _xml.WriteString(Dados.XML.Rows[0]["tpCTe"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("procEmi");
                _xml.WriteString("0"); // 0 - emissão de CT-e com aplicativo do contribuinte
                _xml.WriteEndElement();

                _xml.WriteStartElement("verProc");
                _xml.WriteString(Dados.XML.Rows[0]["verProc"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("cMunEnv");
                _xml.WriteString(Dados.XML.Rows[0]["cMunEnv"].ToString());
                _xml.WriteEndElement();

                _xml.WriteStartElement("xMunEnv");
                _xml.WriteString(Dados.XML.Rows[0]["xMunEnv"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("UFEnv");
                _xml.WriteString(Dados.XML.Rows[0]["UFEnv"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("modal");
                _xml.WriteString(Dados.XML.Rows[0]["modal"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("tpServ");
                _xml.WriteString(Dados.XML.Rows[0]["tpServ"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("cMunIni");
                _xml.WriteString(Dados.XML.Rows[0]["cMunIni"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("xMunIni");
                _xml.WriteString(Dados.XML.Rows[0]["xMunIni"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("UFIni");
                _xml.WriteString(Dados.XML.Rows[0]["UFIni"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("cMunFim");
                _xml.WriteString(Dados.XML.Rows[0]["cMunFim"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("xMunFim");
                _xml.WriteString(Dados.XML.Rows[0]["xMunFim"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("UFFim");
                _xml.WriteString(Dados.XML.Rows[0]["UFFim"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("retira");
                _xml.WriteString(Dados.XML.Rows[0]["retira"].ToString().TrimEnd());
                _xml.WriteEndElement();

                /*_xml.WriteStartElement("xDetRetira"); // marcos@tomazini.org
                _xml.WriteString(Dados.XML.Rows[0]["xDetRetira"].ToString().TrimEnd());
                _xml.WriteEndElement();*/

                if (Dados.XML.Rows[0]["toma"].ToString().TrimEnd() == "4")
                {
                    #region toma04
                    _xml.WriteStartElement("toma4");

                    _xml.WriteStartElement("toma");
                    _xml.WriteString(Dados.XML.Rows[0]["toma"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    string TipoPJ_toma;
                    if (Dados.XML.Rows[0]["toma_CNPJ"].ToString().Replace(".", "").Replace("-", "").Replace("/", "").TrimEnd().Length > 0)
                    {
                        TipoPJ_toma = "J";
                        _xml.WriteStartElement("CNPJ");
                        _xml.WriteString(Dados.XML.Rows[0]["toma_CNPJ"].ToString().Replace(".", "").Replace("-", "").Replace("/", "").TrimEnd());
                        _xml.WriteEndElement();
                    }
                    else
                    {
                        TipoPJ_toma = "F";
                        _xml.WriteStartElement("CPF");
                        _xml.WriteString(Dados.XML.Rows[0]["toma_CPF"].ToString().Replace(".","").Replace("-","").TrimEnd());
                        _xml.WriteEndElement();
                    }

                    _xml.WriteStartElement("IE");
                    if ((Dados.XML.Rows[0]["toma_ie"].ToString().Length > 0) && (TipoPJ_toma == "J"))
                    {
                        _xml.WriteString(Dados.XML.Rows[0]["toma_ie"].ToString().Replace(".", "").TrimEnd());
                    }
                    else
                    {
                        _xml.WriteString("ISENTO");
                    }
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("xNome");
                    _xml.WriteString(Dados.XML.Rows[0]["toma_xNome"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    if (TipoPJ_toma == "J")
                    {
                        if (Dados.XML.Rows[0]["toma_xFant"].ToString().Length > 0)
                        {
                            _xml.WriteStartElement("xFant");
                            _xml.WriteString(Dados.XML.Rows[0]["toma_xFant"].ToString().TrimEnd());
                            _xml.WriteEndElement();
                        }
                    }

                    if (Dados.XML.Rows[0]["endertoma_fone"].ToString().Replace("(", "").Replace(")", "").Replace("-", "").Replace(" ", "").Length > 0)
                    {
                        _xml.WriteStartElement("fone");
                        _xml.WriteString(Dados.XML.Rows[0]["endertoma_fone"].ToString().Replace("(", "").Replace(")", "").Replace("-", "").Replace(" ", "").TrimEnd());
                        _xml.WriteEndElement();
                    }

                    #region enderToma
                    _xml.WriteStartElement("enderToma");

                    _xml.WriteStartElement("xLgr");
                    _xml.WriteString(Dados.XML.Rows[0]["enderToma_xLgr"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("nro");
                    _xml.WriteString(Dados.XML.Rows[0]["enderToma_nro"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    if (Dados.XML.Rows[0]["enderToma_xCpl"].ToString().Length > 0)
                    {
                        _xml.WriteStartElement("xCpl");
                        _xml.WriteString(Dados.XML.Rows[0]["enderToma_xCpl"].ToString().TrimEnd());
                        _xml.WriteEndElement();
                    }

                    _xml.WriteStartElement("xBairro");
                    _xml.WriteString(Dados.XML.Rows[0]["enderToma_xBairro"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("cMun");
                    _xml.WriteString(Dados.XML.Rows[0]["enderToma_cMun"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("xMun");
                    _xml.WriteString(Dados.XML.Rows[0]["enderToma_xMun"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("CEP");
                    _xml.WriteString(Dados.XML.Rows[0]["enderToma_CEP"].ToString().Replace("-", "").TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("UF");
                    _xml.WriteString(Dados.XML.Rows[0]["enderToma_UF"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("cPais");
                    _xml.WriteString("1058");
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("xPais");
                    _xml.WriteString("Brasil");
                    _xml.WriteEndElement();

                    _xml.WriteEndElement();
                    #endregion // enderToma

                    _xml.WriteEndElement();// toma04
                    #endregion
                }
                else
                {
                    #region toma03
                    _xml.WriteStartElement("toma03");

                    _xml.WriteStartElement("toma");
                    _xml.WriteString(Dados.XML.Rows[0]["toma"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteEndElement();// toma03
                    #endregion
                }

                _xml.WriteEndElement();//ide
                //fim ide
                #endregion

                #region compl
                _xml.WriteStartElement("compl");

                /*_xml.WriteStartElement("xCaracAd");
                _xml.WriteString(Dados.XML.Rows[0]["compl_xCaracAd"].ToString());
                _xml.WriteEndElement();

                _xml.WriteStartElement("xCaracSer");
                _xml.WriteString(Dados.XML.Rows[0]["compl_xCaracSer"].ToString());
                _xml.WriteEndElement();*/

                if (Dados.XML.Rows[0]["compl_xEmi"].ToString().Length > 0)
                {
                    _xml.WriteStartElement("xEmi");
                    _xml.WriteString(Dados.XML.Rows[0]["compl_xEmi"].ToString().TrimEnd());
                    _xml.WriteEndElement();
                }

                #region Entrega
                _xml.WriteStartElement("Entrega");

                # region semData
                _xml.WriteStartElement("semData");

                _xml.WriteStartElement("tpPer");
                _xml.WriteString("0"); // Dados.XML.Rows[0]["Entrega_semData_tpPer"].ToString()
                _xml.WriteEndElement();

                _xml.WriteEndElement(); // semData
                #endregion

                # region semHora
                _xml.WriteStartElement("semHora");

                _xml.WriteStartElement("tpHor");
                _xml.WriteString("0");
                _xml.WriteEndElement();

                _xml.WriteEndElement(); // semHora
                #endregion

                _xml.WriteEndElement();// Entrega
                #endregion

                _xml.WriteStartElement("origCalc");
                _xml.WriteString(Dados.XML.Rows[0]["origCalc"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("destCalc");
                _xml.WriteString(Dados.XML.Rows[0]["destCalc"].ToString().TrimEnd());
                _xml.WriteEndElement();

                if (Dados.XML.Rows[0]["xObs"].ToString().Length > 0)
                {
                    _xml.WriteStartElement("xObs");
                    _xml.WriteString(Dados.XML.Rows[0]["xObs"].ToString().TrimEnd());
                    _xml.WriteEndElement();
                }

                _xml.WriteEndElement();// Compl
                #endregion

                #region emit
                _xml.WriteStartElement("emit");

                _xml.WriteStartElement("CNPJ");
                // tratar aqui a retirada da mascara
                _xml.WriteString(Dados.XML.Rows[0]["emit_CNPJ"].ToString().Replace(".", "").Replace("-", "").Replace("/", "").TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("IE");
                _xml.WriteString(Dados.XML.Rows[0]["emit_ie"].ToString().Replace(".","").TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("xNome");
                _xml.WriteString(Dados.XML.Rows[0]["emit_xNome"].ToString().TrimEnd());
                _xml.WriteEndElement();

                if (Dados.XML.Rows[0]["emit_xFant"].ToString().Length > 0)
                {
                    _xml.WriteStartElement("xFant");
                    _xml.WriteString(Dados.XML.Rows[0]["emit_xFant"].ToString().TrimEnd());
                    _xml.WriteEndElement();
                }

                #region enderEmit
                _xml.WriteStartElement("enderEmit");

                _xml.WriteStartElement("xLgr");
                _xml.WriteString(Dados.XML.Rows[0]["enderEmit_xLgr"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("nro");
                _xml.WriteString(Dados.XML.Rows[0]["enderEmit_nro"].ToString().TrimEnd());
                _xml.WriteEndElement();

                if (Dados.XML.Rows[0]["enderEmit_xCpl"].ToString().Length > 0)
                {
                    _xml.WriteStartElement("xCpl");
                    _xml.WriteString(Dados.XML.Rows[0]["enderEmit_xCpl"].ToString().TrimEnd());
                    _xml.WriteEndElement();
                }

                _xml.WriteStartElement("xBairro");
                _xml.WriteString(Dados.XML.Rows[0]["enderEmit_xBairro"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("cMun");
                _xml.WriteString(Dados.XML.Rows[0]["enderEmit_cMun"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("xMun");
                _xml.WriteString(Dados.XML.Rows[0]["enderEmit_xMun"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("CEP");
                _xml.WriteString(Dados.XML.Rows[0]["enderEmit_CEP"].ToString().Replace("-","").TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("UF");
                _xml.WriteString(Dados.XML.Rows[0]["enderEmit_UF"].ToString().TrimEnd());
                _xml.WriteEndElement();

                if (Dados.XML.Rows[0]["enderEmit_fone"].ToString().Replace("(", "").Replace(")", "").Replace("-", "").Replace(" ", "").Length > 0)
                {
                    _xml.WriteStartElement("fone");
                    _xml.WriteString(Dados.XML.Rows[0]["enderEmit_fone"].ToString().Replace("(", "").Replace(")", "").Replace("-", "").Replace(" ", "").TrimEnd());
                    _xml.WriteEndElement();
                }

                _xml.WriteEndElement();
                #endregion

                _xml.WriteEndElement(); // emit
                #endregion

                #region rem
                _xml.WriteStartElement("rem");

                _xml.WriteStartElement("CNPJ");
                _xml.WriteString(Dados.XML.Rows[0]["rem_CNPJ"].ToString().Replace(".", "").Replace("-", "").Replace("/", "").TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("IE");
                _xml.WriteString(Dados.XML.Rows[0]["rem_ie"].ToString().Replace(".","").TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("xNome");
                _xml.WriteString(Dados.XML.Rows[0]["rem_xNome"].ToString().TrimEnd());
                _xml.WriteEndElement();

                if (Dados.XML.Rows[0]["rem_xFant"].ToString().Length > 0)
                {
                    _xml.WriteStartElement("xFant");
                    _xml.WriteString(Dados.XML.Rows[0]["rem_xFant"].ToString().TrimEnd());
                    _xml.WriteEndElement();
                }

                if (Dados.XML.Rows[0]["enderRem_fone"].ToString().Replace("(", "").Replace(")", "").Replace("-", "").Replace(" ", "").Length > 0)
                {
                    _xml.WriteStartElement("fone");
                    _xml.WriteString(Dados.XML.Rows[0]["enderRem_fone"].ToString().Replace("(", "").Replace(")", "").Replace("-", "").Replace(" ", "").TrimEnd());
                    _xml.WriteEndElement();
                }

                #region enderReme
                _xml.WriteStartElement("enderReme");

                _xml.WriteStartElement("xLgr");
                _xml.WriteString(Dados.XML.Rows[0]["enderRem_xLgr"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("nro");
                _xml.WriteString(Dados.XML.Rows[0]["enderRem_nro"].ToString().TrimEnd());
                _xml.WriteEndElement();

                if (Dados.XML.Rows[0]["enderRem_xCpl"].ToString().Length > 0)
                {
                    _xml.WriteStartElement("xCpl");
                    _xml.WriteString(Dados.XML.Rows[0]["enderRem_xCpl"].ToString().TrimEnd());
                    _xml.WriteEndElement();
                }

                _xml.WriteStartElement("xBairro");
                _xml.WriteString(Dados.XML.Rows[0]["enderRem_xBairro"].ToString());
                _xml.WriteEndElement();

                _xml.WriteStartElement("cMun");
                _xml.WriteString(Dados.XML.Rows[0]["enderRem_cMun"].ToString());
                _xml.WriteEndElement();

                _xml.WriteStartElement("xMun");
                _xml.WriteString(Dados.XML.Rows[0]["enderRem_xMun"].ToString());
                _xml.WriteEndElement();

                _xml.WriteStartElement("CEP");
                _xml.WriteString(Dados.XML.Rows[0]["enderRem_CEP"].ToString().Replace("-",""));
                _xml.WriteEndElement();

                _xml.WriteStartElement("UF");
                _xml.WriteString(Dados.XML.Rows[0]["enderRem_UF"].ToString());
                _xml.WriteEndElement();

                _xml.WriteStartElement("cPais");
                _xml.WriteString("1058");
                _xml.WriteEndElement();

                _xml.WriteStartElement("xPais");
                _xml.WriteString("Brasil");
                _xml.WriteEndElement();

                _xml.WriteEndElement();
                #endregion // enderReme

                #region infNF

                foreach (DataRow infNF in Dados.infNF.Rows)
                {

                    if (Dados.XML.Rows[0]["infNF"].ToString() == "0")
                    {
                        _xml.WriteStartElement("infNF");

                        _xml.WriteStartElement("nRoma");
                        _xml.WriteString(infNF["infNF_nRoma"].ToString().TrimEnd());
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("nPed");
                        _xml.WriteString(infNF["infNF_nPed"].ToString().TrimEnd());
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("mod");
                        _xml.WriteString(infNF["infNF_mod"].ToString().TrimEnd());
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("serie");
                        _xml.WriteString(infNF["infNF_serie"].ToString().TrimEnd());
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("nDoc");
                        _xml.WriteString(infNF["infNF_nDoc"].ToString().TrimEnd());
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("dEmi");
                        _xml.WriteString(String.Format("{0:yyyy-MM-dd}", infNF["infNF_dEmi"]));
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("vBC");
                        _xml.WriteString(infNF["infNF_vBC"].ToString().Replace(",", "."));
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("vICMS");
                        _xml.WriteString(infNF["infNF_vICMS"].ToString().Replace(",", "."));
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("vBCST");
                        _xml.WriteString(infNF["infNF_vBCST"].ToString().Replace(",", "."));
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("vST");
                        _xml.WriteString(infNF["infNF_vST"].ToString().Replace(",", "."));
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("vProd");
                        _xml.WriteString(infNF["infNF_vProd"].ToString().Replace(",", "."));
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("vNF");
                        _xml.WriteString(infNF["infNF_vNF"].ToString().Replace(",", "."));
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("nCFOP");
                        _xml.WriteString(infNF["infNF_nCFOP"].ToString().TrimEnd());
                        _xml.WriteEndElement();

                        if (infNF["infNF_nPeso"].ToString().Length > 0)
                        {
                            _xml.WriteStartElement("nPeso");
                            _xml.WriteString(infNF["infNF_nPeso"].ToString().Replace(",", "."));
                            _xml.WriteEndElement();
                        }

                        if (infNF["infNF_PIN"].ToString().Length > 0)
                        {
                            _xml.WriteStartElement("PIN");
                            _xml.WriteString(infNF["infNF_PIN"].ToString().TrimEnd());
                            _xml.WriteEndElement();
                        }

                        _xml.WriteEndElement(); // infNF
                    }
                    else if (Dados.XML.Rows[0]["infNF"].ToString() == "1")
                    {
                        _xml.WriteStartElement("infNFe");

                        _xml.WriteStartElement("chave");
                        _xml.WriteString(infNF["infNFe_chave"].ToString().TrimEnd());
                        _xml.WriteEndElement();

                        if (infNF["infNFe_PIN"].ToString().Length > 0)
                        {
                            _xml.WriteStartElement("PIN");
                            _xml.WriteString(infNF["infNFe_PIN"].ToString().TrimEnd());
                            _xml.WriteEndElement();
                        }

                        _xml.WriteEndElement(); // infNFe                    
                    }
                }
                #endregion

                _xml.WriteEndElement(); // rem
                #endregion

                var i = 0;
                if (i == 1) // nao gerar
                {
                    #region exped
                    _xml.WriteStartElement("exped");

                    _xml.WriteStartElement("CNPJ");
                    _xml.WriteString(Dados.XML.Rows[0]["exped_CNPJ"].ToString().Replace(".", "").Replace("-", "").Replace("/", "").TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("IE");
                    _xml.WriteString(Dados.XML.Rows[0]["exped_IE"].ToString().Replace(".","").TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("xNome");
                    _xml.WriteString(Dados.XML.Rows[0]["exped_xNome"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    /*if (Dados.XML.Rows[0]["exped_xFant"].ToString().Length > 0)
                    {
                        _xml.WriteStartElement("xFant");
                        _xml.WriteString(Dados.XML.Rows[0]["exped_xFant"].ToString());
                        _xml.WriteEndElement();
                    }*/

                    if (Dados.XML.Rows[0]["exped_fone"].ToString().Replace("(", "").Replace(")", "").Replace("-", "").Replace(" ", "").Length > 0)
                    {
                        _xml.WriteStartElement("fone");
                        _xml.WriteString(Dados.XML.Rows[0]["exped_fone"].ToString().Replace("(", "").Replace(")", "").Replace("-", "").Replace(" ", "").TrimEnd());
                        _xml.WriteEndElement();
                    }

                    #region enderExped
                    _xml.WriteStartElement("enderExped");

                    _xml.WriteStartElement("xLgr");
                    _xml.WriteString(Dados.XML.Rows[0]["enderExped_xLgr"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("nro");
                    _xml.WriteString(Dados.XML.Rows[0]["enderExped_nro"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    if (Dados.XML.Rows[0]["enderExped_xCpl"].ToString().Length > 0)
                    {
                        _xml.WriteStartElement("xCpl");
                        _xml.WriteString(Dados.XML.Rows[0]["enderExped_xCpl"].ToString().TrimEnd());
                        _xml.WriteEndElement();
                    }

                    _xml.WriteStartElement("xBairro");
                    _xml.WriteString(Dados.XML.Rows[0]["enderExped_xBairro"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("cMun");
                    _xml.WriteString(Dados.XML.Rows[0]["enderExped_cMun"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("xMun");
                    _xml.WriteString(Dados.XML.Rows[0]["enderExped_xMun"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("CEP");
                    _xml.WriteString(Dados.XML.Rows[0]["enderExped_CEP"].ToString().Replace("-","").TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("UF");
                    _xml.WriteString(Dados.XML.Rows[0]["enderExped_UF"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("cPais");
                    _xml.WriteString("1058");
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("xPais");
                    _xml.WriteString("Brasil");
                    _xml.WriteEndElement();

                    _xml.WriteEndElement();
                    #endregion // enderExped

                    _xml.WriteEndElement(); // exped
                    #endregion

                    #region receb
                    _xml.WriteStartElement("receb");

                    _xml.WriteStartElement("CNPJ");
                    _xml.WriteString(Dados.XML.Rows[0]["receb_CNPJ"].ToString().Replace(".", "").Replace("-", "").Replace("/", "").TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("IE");
                    _xml.WriteString(Dados.XML.Rows[0]["receb_ie"].ToString().Replace("-","").TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("xNome");
                    _xml.WriteString(Dados.XML.Rows[0]["receb_xNome"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    /*if (Dados.XML.Rows[0]["receb_xFant"].ToString().Length > 0)
                    {
                        _xml.WriteStartElement("xFant");
                        _xml.WriteString(Dados.XML.Rows[0]["receb_xFant"].ToString());
                        _xml.WriteEndElement();
                    }*/

                    if (Dados.XML.Rows[0]["receb_fone"].ToString().Replace("(", "").Replace(")", "").Replace("-", "").Replace(" ", "").Length > 0)
                    {
                        _xml.WriteStartElement("fone");
                        _xml.WriteString(Dados.XML.Rows[0]["receb_fone"].ToString().Replace("(", "").Replace(")", "").Replace("-", "").Replace(" ", "").TrimEnd());
                        _xml.WriteEndElement();
                    }

                    #region enderReceb
                    _xml.WriteStartElement("enderReceb");

                    _xml.WriteStartElement("xLgr");
                    _xml.WriteString(Dados.XML.Rows[0]["enderReceb_xLgr"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("nro");
                    _xml.WriteString(Dados.XML.Rows[0]["enderReceb_nro"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    if (Dados.XML.Rows[0]["enderReceb_xCpl"].ToString().Length > 0)
                    {
                        _xml.WriteStartElement("xCpl");
                        _xml.WriteString(Dados.XML.Rows[0]["enderReceb_xCpl"].ToString().TrimEnd());
                        _xml.WriteEndElement();
                    }

                    _xml.WriteStartElement("xBairro");
                    _xml.WriteString(Dados.XML.Rows[0]["enderReceb_xBairro"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("cMun");
                    _xml.WriteString(Dados.XML.Rows[0]["enderReceb_cMun"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("xMun");
                    _xml.WriteString(Dados.XML.Rows[0]["enderReceb_xMun"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("CEP");
                    _xml.WriteString(Dados.XML.Rows[0]["enderReceb_CEP"].ToString().Replace("-","").TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("UF");
                    _xml.WriteString(Dados.XML.Rows[0]["enderReceb_UF"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("cPais");
                    _xml.WriteString("1058");
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("xPais");
                    _xml.WriteString("Brasil");
                    _xml.WriteEndElement();

                    _xml.WriteEndElement();
                    #endregion // enderReceb

                    _xml.WriteEndElement(); // receb
                    #endregion
                }

                #region dest
                _xml.WriteStartElement("dest");

                string TipoPJ_Dest;
                if (Dados.XML.Rows[0]["dest_CNPJ"].ToString().Replace(".", "").Replace("-", "").Replace("/", "").Length > 0)
                {
                    TipoPJ_Dest = "J";
                    _xml.WriteStartElement("CNPJ");
                    _xml.WriteString(Dados.XML.Rows[0]["dest_CNPJ"].ToString().Replace(".", "").Replace("-", "").Replace("/", "").TrimEnd());
                    _xml.WriteEndElement();
                }
                else
                {
                    TipoPJ_Dest = "F";
                    _xml.WriteStartElement("CPF");
                    _xml.WriteString(Dados.XML.Rows[0]["dest_CPF"].ToString().Replace(".","").Replace("-","").TrimEnd());
                    _xml.WriteEndElement();
                }

                _xml.WriteStartElement("IE");
                if ((Dados.XML.Rows[0]["dest_ie"].ToString().Replace(".","").Length > 0) && (TipoPJ_Dest == "J"))
                {
                    _xml.WriteString(Dados.XML.Rows[0]["dest_ie"].ToString().Replace(".", "").TrimEnd());
                }
                else
                {
                    _xml.WriteString("ISENTO");
                }
                _xml.WriteEndElement();

                _xml.WriteStartElement("xNome");
                _xml.WriteString(Dados.XML.Rows[0]["dest_xNome"].ToString().TrimEnd());
                _xml.WriteEndElement();

                /*if (TipoPJ_Dest == "J")
                {
                    if (Dados.XML.Rows[0]["dest_xFant"].ToString().Length > 0)
                    {
                        _xml.WriteStartElement("xFant");
                        _xml.WriteString(Dados.XML.Rows[0]["dest_xFant"].ToString());
                        _xml.WriteEndElement();
                    }
                }*/

                if (Dados.XML.Rows[0]["dest_fone"].ToString().Replace("(", "").Replace(")", "").Replace("-", "").Replace(" ", "").Length > 0)
                {
                    _xml.WriteStartElement("fone");
                    _xml.WriteString(Dados.XML.Rows[0]["dest_fone"].ToString().Replace("(", "").Replace(")", "").Replace("-", "").Replace(" ", "").TrimEnd());
                    _xml.WriteEndElement();
                }

                #region enderDest
                _xml.WriteStartElement("enderDest");

                _xml.WriteStartElement("xLgr");
                _xml.WriteString(Dados.XML.Rows[0]["enderDest_xLgr"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("nro");
                _xml.WriteString(Dados.XML.Rows[0]["enderDest_nro"].ToString().TrimEnd());
                _xml.WriteEndElement();

                if (Dados.XML.Rows[0]["enderDest_xCpl"].ToString().Length > 0)
                {
                    _xml.WriteStartElement("xCpl");
                    _xml.WriteString(Dados.XML.Rows[0]["enderDest_xCpl"].ToString().TrimEnd());
                    _xml.WriteEndElement();
                }

                _xml.WriteStartElement("xBairro");
                _xml.WriteString(Dados.XML.Rows[0]["enderDest_xBairro"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("cMun");
                _xml.WriteString(Dados.XML.Rows[0]["enderDest_cMun"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("xMun");
                _xml.WriteString(Dados.XML.Rows[0]["enderDest_xMun"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("CEP");
                _xml.WriteString(Dados.XML.Rows[0]["enderDest_CEP"].ToString().Replace("-","").TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("UF");
                _xml.WriteString(Dados.XML.Rows[0]["enderDest_UF"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("cPais");
                _xml.WriteString("1058");
                _xml.WriteEndElement();

                _xml.WriteStartElement("xPais");
                _xml.WriteString("Brasil");
                _xml.WriteEndElement();

                _xml.WriteEndElement();
                #endregion // enderDest

                _xml.WriteEndElement(); // dest
                #endregion

                #region vPrest
                _xml.WriteStartElement("vPrest");

                _xml.WriteStartElement("vTPrest");
                _xml.WriteString(Dados.XML.Rows[0]["vPrest_vTPrest"].ToString().Replace(",", "."));
                _xml.WriteEndElement();

                _xml.WriteStartElement("vRec");
                _xml.WriteString(Dados.XML.Rows[0]["vPrest_vRec"].ToString().Replace(",", "."));
                _xml.WriteEndElement();

                #region Comp
                foreach (DataRow Comp in Dados.Comp.Rows)
                {
                    _xml.WriteStartElement("Comp");

                    _xml.WriteStartElement("xNome");
                    _xml.WriteString(Comp["xNome"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("vComp");
                    _xml.WriteString(Comp["vComp"].ToString().Replace(",", "."));
                    _xml.WriteEndElement();

                    _xml.WriteEndElement();
                }
                #endregion

                _xml.WriteEndElement();
                #endregion

                #region imp
                _xml.WriteStartElement("imp");

                #region ICMS
                _xml.WriteStartElement("ICMS");
                string CST = Dados.XML.Rows[0]["icms_CST"].ToString().TrimEnd();
                switch (CST)
                {
                    case "00":
                        #region ICMS00
                        _xml.WriteStartElement("ICMS00");

                        _xml.WriteStartElement("CST");
                        _xml.WriteString(Dados.XML.Rows[0]["icms_CST"].ToString().TrimEnd());
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("vBC");
                        _xml.WriteString(Dados.XML.Rows[0]["icms_vBC"].ToString().Replace(",", "."));
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("pICMS");
                        _xml.WriteString(Dados.XML.Rows[0]["icms_pICMS"].ToString().Replace(",", "."));
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("vICMS");
                        _xml.WriteString(Dados.XML.Rows[0]["icms_vICMS"].ToString().Replace(",", "."));
                        _xml.WriteEndElement();

                        _xml.WriteEndElement();
                        #endregion
                        break;

                    case "20":
                        #region ICMS20
                        _xml.WriteStartElement("ICMS20");

                        _xml.WriteStartElement("CST");
                        _xml.WriteString(Dados.XML.Rows[0]["icms_CST"].ToString().Replace(",", "."));
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("pRedBC");
                        _xml.WriteString(Dados.XML.Rows[0]["icms_pRedBC"].ToString().Replace(",", "."));
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("vBC");
                        _xml.WriteString(Dados.XML.Rows[0]["icms_vBC"].ToString().Replace(",", "."));
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("pICMS");
                        _xml.WriteString(Dados.XML.Rows[0]["icms_pICMS"].ToString().Replace(",", "."));
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("vICMS");
                        _xml.WriteString(Dados.XML.Rows[0]["icms_vICMS"].ToString().Replace(",", "."));
                        _xml.WriteEndElement();

                        _xml.WriteEndElement();
                        #endregion
                        break;

                    case "40":
                    case "41":
                    case "51":
                        #region ICMS45
                        _xml.WriteStartElement("ICMS45");

                        _xml.WriteStartElement("CST");
                        _xml.WriteString(Dados.XML.Rows[0]["icms_CST"].ToString().TrimEnd());
                        _xml.WriteEndElement();

                        _xml.WriteEndElement();
                        #endregion
                        break;

                    case "60":
                        #region ICMS60
                        _xml.WriteStartElement("ICMS60");

                        _xml.WriteStartElement("CST");
                        _xml.WriteString(Dados.XML.Rows[0]["icms_CST"].ToString().Replace(",", "."));
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("vBCSTRet");
                        _xml.WriteString(Dados.XML.Rows[0]["icms_vBC"].ToString().Replace(",", "."));
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("vICMSSTRet");
                        _xml.WriteString(Dados.XML.Rows[0]["icms_vICMS"].ToString().Replace(",", "."));
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("pICMSSTRet");
                        _xml.WriteString(Dados.XML.Rows[0]["icms_pICMS"].ToString().Replace(",", "."));
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("vCred");
                        _xml.WriteString(Dados.XML.Rows[0]["icms_vCred"].ToString().Replace(",", "."));
                        _xml.WriteEndElement();

                        _xml.WriteEndElement();
                        #endregion
                        break;

                    case "90":
                        #region ICMS90
                        _xml.WriteStartElement("ICMS90");

                        _xml.WriteStartElement("CST");
                        _xml.WriteString(Dados.XML.Rows[0]["icms_CST"].ToString().Replace(",", "."));
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("pRedBC");
                        _xml.WriteString(Dados.XML.Rows[0]["icms_pRedBC"].ToString().Replace(",", "."));
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("vBC");
                        _xml.WriteString(Dados.XML.Rows[0]["icms_vBC"].ToString().Replace(",", "."));
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("pICMS");
                        _xml.WriteString(Dados.XML.Rows[0]["icms_pICMS"].ToString().Replace(",", "."));
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("vICMS");
                        _xml.WriteString(Dados.XML.Rows[0]["icms_vICMS"].ToString().Replace(",", "."));
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("vCred");
                        _xml.WriteString(Dados.XML.Rows[0]["icms_vCred"].ToString().Replace(",", "."));
                        _xml.WriteEndElement();

                        _xml.WriteEndElement();
                        #endregion
                        break;

                    case "91":
                        #region ICMSOutraUF - CST 91
                        _xml.WriteStartElement("ICMSOutraUF");

                        _xml.WriteStartElement("CST");
                        _xml.WriteString(Dados.XML.Rows[0]["icms_CST"].ToString().Replace(",", "."));
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("pRedBCOutraUF");
                        _xml.WriteString(Dados.XML.Rows[0]["icms_pRedBC"].ToString().Replace(",", "."));
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("vBCOutraUF");
                        _xml.WriteString(Dados.XML.Rows[0]["icms_vBC"].ToString().Replace(",", "."));
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("pICMSOutraUF");
                        _xml.WriteString(Dados.XML.Rows[0]["icms_pICMS"].ToString().Replace(",", "."));
                        _xml.WriteEndElement();

                        _xml.WriteStartElement("vICMSOutraUF");
                        _xml.WriteString(Dados.XML.Rows[0]["icms_vICMS"].ToString().Replace(",", "."));
                        _xml.WriteEndElement();

                        _xml.WriteEndElement();
                        #endregion
                        break;

                    case "SN":
                        #region ICMSSN
                        _xml.WriteStartElement("ICMSSN");

                        _xml.WriteStartElement("indSN");
                        _xml.WriteString("1"); // Indica se o contribuinte é Simples Nacional 1=Sim
                        _xml.WriteEndElement();

                        _xml.WriteEndElement();
                        #endregion
                        break;
                }

                _xml.WriteEndElement();
                #endregion

                _xml.WriteEndElement();
                #endregion

                #region infCTeNorm
                _xml.WriteStartElement("infCTeNorm");

                #region infCarga
                _xml.WriteStartElement("infCarga");

                _xml.WriteStartElement("vCarga");
                _xml.WriteString(Dados.XML.Rows[0]["infCarga_vCarga"].ToString().Replace(",", "."));
                _xml.WriteEndElement();

                _xml.WriteStartElement("proPred");
                _xml.WriteString(Dados.XML.Rows[0]["infCarga_proPred"].ToString().TrimEnd());
                _xml.WriteEndElement();

                if (Dados.XML.Rows[0]["infCarga_xOutCat"].ToString().Length > 0)
                {
                    _xml.WriteStartElement("xOutCat");
                    _xml.WriteString(Dados.XML.Rows[0]["infCarga_xOutCat"].ToString().TrimEnd());
                    _xml.WriteEndElement();
                }

                #region infQ
                foreach (DataRow infQ in Dados.infQ.Rows)
                {
                    _xml.WriteStartElement("infQ");

                    _xml.WriteStartElement("cUnid");
                    _xml.WriteString(infQ["cUnid"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("tpMed");
                    _xml.WriteString(infQ["tpMed"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("qCarga");
                    _xml.WriteString(infQ["qCarga"].ToString().Replace(",", "."));
                    _xml.WriteEndElement();

                    _xml.WriteEndElement();
                }
                #endregion // infQ

                _xml.WriteEndElement();
                #endregion // infCarga

                #region seg
                if (Dados.XML.Rows[0]["respSeg"].ToString().Length > 0)
                {
                    _xml.WriteStartElement("seg");

                    _xml.WriteStartElement("respSeg");
                    _xml.WriteString(Dados.XML.Rows[0]["respSeg"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteEndElement();
                }
                #endregion

                #region infModal
                _xml.WriteStartElement("infModal");

                _xml.WriteStartAttribute("versaoModal");
                _xml.WriteString("1.04");
                _xml.WriteEndAttribute();

                #region rodo
                _xml.WriteStartElement("rodo");

                _xml.WriteStartElement("RNTRC");
                _xml.WriteString(Dados.XML.Rows[0]["rodo_RNTRC"].ToString().TrimEnd());
                _xml.WriteEndElement();

                _xml.WriteStartElement("dPrev");
                Console.WriteLine(String.Format("{0:yyyy-MM-dd}", Dados.XML.Rows[0]["rodo_dPrev"]));
                _xml.WriteString(String.Format("{0:yyyy-MM-dd}", Dados.XML.Rows[0]["rodo_dPrev"]));
                _xml.WriteEndElement();

                _xml.WriteStartElement("lota");
                _xml.WriteString(Dados.XML.Rows[0]["rodo_lota"].ToString().TrimEnd());
                _xml.WriteEndElement();

                if (Dados.XML.Rows[0]["rodo_CIOT"].ToString().Length > 0)
                {
                    _xml.WriteStartElement("CIOT");
                    _xml.WriteString(Dados.XML.Rows[0]["rodo_CIOT"].ToString().TrimEnd());
                    _xml.WriteEndElement();
                }

                #region veic
                if (Dados.XML.Rows[0]["id_veic"].ToString().Length > 0)
                {
                    _xml.WriteStartElement("veic");

                    /*_xml.WriteStartElement("cInt");
                    _xml.WriteString(Dados.XML.Rows[0]["cInt"].ToString());
                    _xml.WriteEndElement();*/

                    _xml.WriteStartElement("RENAVAM");
                    _xml.WriteString(Dados.XML.Rows[0]["RENAVAM"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("placa");
                    _xml.WriteString(Dados.XML.Rows[0]["placa"].ToString().Replace("-","").TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("tara");
                    _xml.WriteString(Dados.XML.Rows[0]["tara"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("capKG");
                    _xml.WriteString(Dados.XML.Rows[0]["capKG"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("capM3");
                    _xml.WriteString(Dados.XML.Rows[0]["capM3"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("tpProp");
                    _xml.WriteString(Dados.XML.Rows[0]["tpProp"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("tpVeic");
                    _xml.WriteString(Dados.XML.Rows[0]["tpVeic"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("tpRod");
                    _xml.WriteString(Dados.XML.Rows[0]["tpRod"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("tpCar");
                    _xml.WriteString(Dados.XML.Rows[0]["tpCar"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("UF");
                    _xml.WriteString(Dados.XML.Rows[0]["veicUF"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteEndElement();
                }
                #endregion // veic

                #region moto
                foreach (DataRow moto in Dados.motorista.Rows)
                {
                    _xml.WriteStartElement("moto");

                    _xml.WriteStartElement("xNome");
                    _xml.WriteString(moto["xNome"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("CPF");
                    _xml.WriteString(moto["CPF"].ToString().Replace(".","").Replace("-","").TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteEndElement();
                }
                #endregion // moto

                _xml.WriteEndElement();
                #endregion // rodo

                _xml.WriteEndElement();
                #endregion // infModal

                #region cobr
                if (Dados.XML.Rows[0]["fat_nFat"].ToString().Length > 0)
                {
                    _xml.WriteStartElement("cobr");

                    #region fat
                    _xml.WriteStartElement("fat");

                    _xml.WriteStartElement("nFat");
                    _xml.WriteString(Dados.XML.Rows[0]["fat_nFat"].ToString().TrimEnd());
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("vOrig");
                    _xml.WriteString(Dados.XML.Rows[0]["fat_vOrig"].ToString().Replace(",", "."));
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("vDesc");
                    _xml.WriteString(Dados.XML.Rows[0]["fat_vDesc"].ToString().Replace(",", "."));
                    _xml.WriteEndElement();

                    _xml.WriteStartElement("vLiq");
                    _xml.WriteString(Dados.XML.Rows[0]["fat_vLiq"].ToString().Replace(",", "."));
                    _xml.WriteEndElement();

                    _xml.WriteEndElement();
                    #endregion // fat

                    _xml.WriteEndElement();
                }
                #endregion // cobr

                _xml.WriteEndElement();
                #endregion

                //END InfCTe
                _xml.WriteEndElement();//InfCTe
                #endregion
                //end nfe
                _xml.WriteEndElement();//CTe
                #endregion

                //_xml.Flush();
                //Constantes.XML_Sem_Assinatura = swsw.ToString();

                XmlDocument xdoc = new XmlDocument();

                string XMLFormatado = RemoveCaracterEspeciais(sw.ToString());
                XMLFormatado = XMLFormatado.Replace("<?xml version=\"1.0\" encoding=\"utf-16\"?>", "");

                xdoc.LoadXml(XMLFormatado);

                Constantes.XML_Sem_Assinatura = xdoc;

                _xml.Close();

                //MessageBox.Show("Arquivo exportado\nSucesso"); // marcos@tomazini.org
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao Gerar XML\n" + ex.ToString());
            }
        }
    }
}
