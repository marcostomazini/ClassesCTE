using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Security.Cryptography.X509Certificates;

namespace EmissorCTe.CS
{
    class Certificado
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

        //
        /*
         *         Buscar : 0 - Busca realizada com sucesso
         *                  1 - Nenhum certificado escolhido
         *                  2 - Nenhum certificado válido foi encontrado com o nome informado: + Nome
         *                  3 - Exceção na busca por nome
         *                  4 - Nenhum certificado válido foi encontrado com o número de série informado: + NroSerie
         *                  5 - Nenhum certificado válido foi encontrado com o número de série informado: + NroSerie
         *                  6 - Exceção na busca por número de série
         */

        public X509Certificate2 BuscaNome(string pNome)
        {
            //pNome = "VAGNER MIQUILINO";
            resultado = 0;
            msgResultado = "Busca realizada com sucesso.";           

            X509Certificate2 oX509Certificado = new X509Certificate2();
            try
            {
                X509Store oStore = new X509Store("MY", StoreLocation.CurrentUser);
                oStore.Open(OpenFlags.ReadOnly | OpenFlags.OpenExistingOnly);
                X509Certificate2Collection oCollection = (X509Certificate2Collection)oStore.Certificates;
                X509Certificate2Collection oCollection1 = (X509Certificate2Collection)oCollection.Find(X509FindType.FindByTimeValid, DateTime.Now, false);
                X509Certificate2Collection oCollection2 = (X509Certificate2Collection)oCollection.Find(X509FindType.FindByKeyUsage, X509KeyUsageFlags.DigitalSignature, false);

                if (pNome == "")
                {
                    X509Certificate2Collection oCollection3 = X509Certificate2UI.SelectFromCollection(oCollection2, "Certificado(s) Digital(is) disponível(is)", "Selecione o Certificado Digital para uso no aplicativo", X509SelectionFlag.SingleSelection);
                    if (oCollection3.Count == 0)
                    {
                        resultado = 1;
                        msgResultado = "Nenhum certificado digital foi selecionado ou o certificado selecionado está com problemas.";
                        oX509Certificado.Reset();
                    }
                    else
                    {
                        oX509Certificado = oCollection3[0];
                    }
                }
                else
                {
                    //System.Security.Cryptography.X509Certificates.X509Certificate2Collection oCollection3 = (X509Certificate2Collection)oCollection2.Find(X509FindType.FindByThumbprint, pNome, false);
                    System.Security.Cryptography.X509Certificates.X509Certificate2Collection oCollection3 = (X509Certificate2Collection)oCollection2.Find(X509FindType.FindBySubjectName, pNome, false);
                    if (oCollection3.Count == 0)
                    {
                        resultado = 2;
                        msgResultado = "Nenhum certificado válido foi encontrado com o nome informado: " + pNome;
                        oX509Certificado.Reset();
                    }
                    else
                    {
                        oX509Certificado = oCollection3[0];
                    }
                }
                oStore.Close();
                return oX509Certificado;
            }
            catch (System.Exception ex)
            {
                resultado = 3;
                msgResultado = ex.Message;
                return oX509Certificado;
            }
        }
    }
}
