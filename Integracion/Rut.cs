using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Integracion
{
    public class Rut
    {
        public string RutCompleto
        {
            get
            {
                string rut = Numero.ToString();
                string conPuntos = "";
                int contador = 0;
                string[] cadena = rut.Split();
                for (int indice = rut.Length; indice > 0; indice--)
                {
                    conPuntos = rut.Substring(indice - 1, 1) + conPuntos;
                    contador++;
                    if (contador >= 3)
                    {
                        conPuntos = "." + conPuntos;
                        contador = 0;
                    }
                }
                return conPuntos + '-' + DV;
            }
        }
        public int Numero { get; set; }
        public char DV { get; set; }

        public Rut() { }
        public Rut(int numero)
        {
            Numero = numero;
            DV = char.Parse(DigitoVerificador(numero));
        }

        public static Rut Parse(string s)
        {
            if (String.IsNullOrEmpty(s)) throw new ArgumentNullException("s");

            Rut result = new Rut();
            s = s.Trim().Replace("-", "");
            s = s.Trim().Replace(".", "");
            string rutTitularString = s.Substring(0, s.Length - 1);
            result.Numero = int.Parse(rutTitularString);
            result.DV = char.Parse(s.Substring(s.Length - 1).ToUpper());

            return result;
        }

        public override string ToString()
        {
            if (!EsValido) return String.Empty;
            return String.Format("{0}{1}", Numero, DV);
        }

        public string ToStringConGuion()
        {
            if (!EsValido) return String.Empty;
            return String.Format("{0}-{1}", Numero.ToString("N0"), DV);
        }

        public static bool TryParse(string s, out Rut rut)
        {
            try
            {
                rut = Rut.Parse(s);
                return true;
            }
            catch (Exception ex)
            {
                rut = null;
                return false;
            }
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="rutConDV">Rut en formato BSS (Sin puntos ni guion, con DV al final)</param>
        public Rut(string rutConDV)
        {
            rutConDV = rutConDV.Trim();
            rutConDV = rutConDV.Replace(".", "");
            string rutTitularString = rutConDV.Substring(0, rutConDV.Length - 1);
            try
            {
                Numero = int.Parse(rutTitularString);
            }
            catch (Exception e)
            {
                throw new System.FormatException(string.Format("El valor {0} no puede convertirse a rut", rutConDV), e);
            }
            DV = rutConDV.Substring(rutConDV.Length - 1).ToUpper().ToCharArray()[0];
        }

        /// <summary>
        /// Revisa la validez del RUT utilizando el DV.
        /// </summary>
        public bool EsValido
        {
            get
            {
                return (VerificaRut(this.Numero, this.DV));
            }
        }

        /// <summary>
        /// Verifica que un rut sea valido
        /// </summary>
        /// <param name="Rut"></param>
        /// <param name="DV"></param>
        /// <returns></returns>
        static public bool VerificaRut(int Rut, char DV)
        {
            // Chequeos que se hacen: 

            if ((Rut.ToString().Length > 11) || (DV.ToString().ToUpper() != DigitoVerificador(Rut)))
                return false;
            else
                return true;
        }

        /// <summary>
        /// Calcula el digito verificador de un RUT
        /// </summary>
        /// <param name="rut"></param>
        /// <returns>Digito verificador del rut</returns>
        static public string DigitoVerificador(int rut)
        {
            int Digito;
            int Contador;
            int Multiplo;
            int Acumulador;
            string RutDigito;

            Contador = 2;
            Acumulador = 0;

            while (rut != 0)
            {
                Multiplo = (rut % 10) * Contador;
                Acumulador = Acumulador + Multiplo;
                rut = rut / 10;
                Contador = Contador + 1;
                if (Contador == 8)
                {
                    Contador = 2;
                }

            }

            Digito = 11 - (Acumulador % 11);
            RutDigito = Digito.ToString().Trim();
            if (Digito == 10)
            {
                RutDigito = "K";
            }
            if (Digito == 11)
            {
                RutDigito = "0";
            }
            return (RutDigito);
        }
    }
}
