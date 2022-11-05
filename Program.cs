using System.Net.Mail;
using System.Net;
using System.Xml.Linq;
using KolicinaBarkodTEST;

namespace KontrolaKolicina
{

    /// <summary>
    /// 
    /// 
    /// Ovaj program treba da resi problem kada se u kolicine skenira barkod artikla,
    /// 
    /// To izgleda tako sto se u kolicinu ubaci barkod i onda se ta kolicina mnozi u odredjene atribute
    /// 
    /// IZR_VREDNOST = NA_IZR_MPV = NA_IZR_MPV_BPOPDOK  -    sto je NA_IZR_MPC * KOLICINA
    /// NA_IZR_PV = NA_IZR_PV_BPOPDOK                   -    sto je NA_IZR_PC * KOLICINA
    /// 
    /// Jedan artikal bude u pozitivnoj kolicini a jedan u negativnoj (to je jer kod nas kada se
    /// stavka sa racuna obrise, stvara se ista kolicina stavke samo u negativnoj protiv-vrednosti
    /// 
    /// Ovaj program prodje kroz XML dokument, i treba da nadje POSTAVKU (stavku) gde je kolicina veca
    /// od 9999 ili manja od -9999 (to bi trebalo da su barkodovi jer kod nas se nista ne prodaje u tim kolicinama)
    /// Onda izmeni kolicine na jedan i minus jedan a ostale vrednosti po NA_IZR_MPC i NA_IZR_PC
    /// 
    /// </summary>
    /// <param name="args"></param>
    /// 




    internal class Program
    {
        static void Main(string[] args)
        {

            //da mogu da dobijem vreme za log
            DateTime now = DateTime.Now;


            //pretraga svih kontrola na 43ci
            string[] kontrole = Directory.GetDirectories(@"\\LOKACIJAKONTROLA"
                        , "Kontrola*", SearchOption.TopDirectoryOnly);



            //info za slanje maila
            string body = "Izvrseno je:\n";

            //Putanja gde prebacujemo racune koje smo pronasli da su losi
            string pronadjeniRacuni = @"C:\\Lilly\KontrolaKolicina\PronadjeniRacuni\" + now.ToString("dd.MM.yyyy");
            //Putanja gde prebacujemo racune koje smo pronasli da su prazni
            string prazniRacuni = @"C:\\Lilly\KontrolaKolicina\PrazniRacuni\" + now.ToString("dd.MM.yyyy");

            //OVDE UPISUJES GDE SE CUVA RACUN
            string ispravljeniRacuni = @"C:\\Lilly\KontrolaKolicina\IspravljeniRacuni\" + now.ToString("dd.MM.yyyy");

            //Pravljenje foldera za logovanje
            string dirLogFile = @"C:\\Lilly\KontrolaKolicina\LogFiles\" + now.ToString("dd.MM.yyyy");


            #region "Pravljenje foldera za skladistenje (Log,Ispravljeni,Prazni,Pronadjeni)"

            if (!Directory.Exists(dirLogFile))
            {
                Directory.CreateDirectory(dirLogFile);
            }
            else//ovo je ako se program pokrene vise puta na dan (da se doda na folder i sat min sek)
            {
                dirLogFile += "_" + now.ToString("HHmmss");
                Directory.CreateDirectory(dirLogFile);

            }
            #region "Folderi"
            if (!Directory.Exists(pronadjeniRacuni))
            {
                Directory.CreateDirectory(pronadjeniRacuni);
            }
            else//ovo je ako se program pokrene vise puta na dan (da se doda na folder i sat min sek)
            {
                pronadjeniRacuni += "_" + now.ToString("HHmmss");
                Directory.CreateDirectory(pronadjeniRacuni);

            }


            if (!Directory.Exists(ispravljeniRacuni))
            {
                Directory.CreateDirectory(ispravljeniRacuni);
            }
            else//ovo je ako se program pokrene vise puta na dan (da se doda na folder i sat min sek)
            {
                ispravljeniRacuni += "_" + now.ToString("HHmmss");
                Directory.CreateDirectory(ispravljeniRacuni);

            }
            #endregion

            #endregion



            List<string> napakeKontrole = new List<string>();
            //prolazak kroz array kontrole i dodaje se u listu sa nastavkom do Napake
            foreach (string s in kontrole)
            {
                napakeKontrole.Add(s + @"\Queue\Napake");
            }

            //da izbrojimo ispravljen broj artiakala
            int count = 0;

            List<string> losiRacuni = new List<string>();

            //prolazak kroz listu da nadjemo putanje do Napake
            foreach (string kontrola in napakeKontrole)
            {
                //UZIMANJE FILEOVA IZ FOLDERA
                //Uzima sve filepath iz foldera Napake, koji pocinju sa "PAR" a zarvsavaju se sa ".xml"
                //i stavlja ih u array
                string[] filePaths = Directory.GetFiles(kontrola
                        , "PAR*.xml", SearchOption.TopDirectoryOnly);

                //ako nema racuna 
                if (filePaths == null || filePaths.Length == 0)
                {
                    using (StreamWriter _log = new StreamWriter(dirLogFile + "\\" + now.ToString("dd_MM_yyyy") + "_" + "SveOkLog.txt"
                        , true))
                    {
                        _log.WriteLine(now.ToString("G") + "\t"
                                + "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");

                        _log.WriteLine(now.ToString("G") + "\t"
                            + $"Nema racuna u {kontrola}");

                        _log.WriteLine(now.ToString("G") + "\t"
                                    + "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
                    }
                    continue;
                }


                //prolazak kroz listu filePaths (xmlFile je putanja sa nazivom racuna)
                foreach (string xmlFile in filePaths)
                {


                    //Ovde uzimamo putanju do file-a i onda koristimo Substring
                    //da nadjemo naziv file-a. Ovo planiram da koristim kada budem sacuvao file
                    string nazivRacuna = xmlFile.Substring(xmlFile.LastIndexOf('\\') + 1);

                    //ovde kreiramo objekat klase PodaciRacuna odakle mozemo da izvucemo vise informacija
                    PodaciRacuna infoRacun = new PodaciRacuna(nazivRacuna);

                    try // ovaj try koristimo da izbegnemo Exception gde postoje racuni bez vrednosti u njima
                    {
                        //ucitava xmlFile u program
                        XDocument xDoc = XDocument.Load(xmlFile);


                        //Nalazi Node POSTAVKA
                        var postavke = xDoc.Root.Descendants("POSTAVKA");

                        //provera da li postavke imaju podakte
                        if (postavke != null)
                        {
                            //Prolazi kroz sve postavke
                            foreach (var postavka in postavke)
                            {
                                //Uzima atribut KOLICINA iz postavke i pokusava da ga konvertuje u double
                                double kolicina;
                                bool res = double.TryParse(postavka.Attribute("KOLICINA").Value, out kolicina);
                                //ako je konverzija u double uspela
                                if (res == true)
                                {
                                    //ako je kolicina pogresna, dodaj u body mail-a podatke racuna
                                    if (kolicina > 9999 || kolicina < -9999)
                                    {


                                        // da uracunamo artikal koji je ispravljne
                                        count++;
                                        if (!losiRacuni.Contains(nazivRacuna))
                                        {
                                            losiRacuni.Add(nazivRacuna);
                                        }

                                        //ako je kolicina veca od 9999 ili negativna
                                        if (kolicina > 9999)
                                        {
                                            using (StreamWriter _log = new StreamWriter(dirLogFile + "\\" + now.ToString("dd_MM_yyyy") + "_" + "PromeneLog.txt", true))
                                            {
                                                //logovanje
                                                _log.WriteLine("\n" + now.ToString("G") + "\t"
                                                    + $"Kontrola je: {kontrola}");

                                                _log.WriteLine(now.ToString("G") + "\t"
                                                    + $"Naziv racuna: {nazivRacuna}");


                                            }

                                            //Ovde menjamo vrednosti i pisemo log
                                            LogIzmenaPozitivne(postavka, dirLogFile, now, infoRacun);


                                        }
                                        //da li je kolicina manja od -9999, ovo je da nadjemo onaj gde je obrisan artikal
                                        if (kolicina < -9999)
                                        {


                                            //Ovde menjamo vrednosti i pisemo log
                                            LogIzmenaNegativne(postavka, dirLogFile, now);



                                        }


                                        
                                        xDoc.Save(ispravljeniRacuni + "\\" + nazivRacuna);



                                        //Prebaci racun u PronadjeniRacuni
                                        if (!File.Exists(pronadjeniRacuni + "\\" + nazivRacuna))
                                        {
                                            File.Move(xmlFile, pronadjeniRacuni + "\\" + nazivRacuna, true);
                                        }



                                    }


                                }
                                else
                                {

                                    using (StreamWriter _log = new StreamWriter(dirLogFile + "\\" + now.ToString("dd_MM_yyyy") + "_" + "PromeneLog.txt", true))
                                    {
                                        _log.WriteLine("\n");

                                        _log.WriteLine(now.ToString("G") + "\t"
                                            + "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");

                                        _log.WriteLine(now.ToString("G") + "\t"
                                            + $"Racun: {nazivRacuna} je prazan");

                                        _log.WriteLine(now.ToString("G") + "\t"
                                            + "PARSIRANJE KOLICINE NIJE USPELO");

                                        _log.WriteLine(now.ToString("G") + "\t"
                                            + "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");


                                        _log.WriteLine("\n");
                                    }
                                }
                            }
                        }
                        else
                        {
                            using (StreamWriter _log = new StreamWriter(dirLogFile + "\\" + now.ToString("dd_MM_yyyy") + "_" + "PromeneLog.txt", true))
                            {
                                _log.WriteLine("\n");

                                _log.WriteLine(now.ToString("G") + "\t"
                                    + "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
                                _log.WriteLine(now.ToString("G") + "\t"
                                + $"Racun: {nazivRacuna} je prazan");

                                _log.WriteLine(now.ToString("G") + "\t"
                                    + "NISU PRONADJENE STAVKE NA RACUNU");

                                _log.WriteLine(now.ToString("G") + "\t"
                                    + "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");


                                _log.WriteLine("\n");
                            }
                        }

                    }
                    catch (System.Xml.XmlException)
                    {

                        if (!Directory.Exists(prazniRacuni))
                        {
                            Directory.CreateDirectory(prazniRacuni);
                        }
                        else//ovo je ako se program pokrene vise puta na dan (da se doda na folder i sat min sek)
                        {
                            prazniRacuni += "_" + now.ToString("HHmmss");
                            Directory.CreateDirectory(prazniRacuni);

                        }


                        using (StreamWriter _log = new StreamWriter(dirLogFile + "\\" + now.ToString("dd_MM_yyyy") + "_" + "PromeneLog.txt", true))
                        {

                            _log.WriteLine("\n");

                            _log.WriteLine(now.ToString("G") + "\t"
                                + "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");

                            _log.WriteLine(now.ToString("G") + "\t"
                                + $"Racun: {nazivRacuna} je prazan");
                            _log.WriteLine(now.ToString("G") + "\t"
                                + $"Racun: {nazivRacuna} je prebacen u (172.16.130.74) {prazniRacuni}");

                            _log.WriteLine(now.ToString("G") + "\t"
                                        + "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");

                            _log.WriteLine("\n");
                        }


                        //Prebaci racun u PrazniRacuni
                        if (!File.Exists(prazniRacuni + "\\" + nazivRacuna))
                        {
                            File.Move(xmlFile, prazniRacuni + "\\" + nazivRacuna, true);
                        }




                        continue;
                    }//catch

                    // }//StreamWriter

                }//xmlFile foreach
            }//napakeKontrola foreach

            //provera da li smo nasli neke racune, ako da saljemo mail
            if (losiRacuni.Count() != 0)
            {
                body += $"Broj losih racuna je : {losiRacuni.Count()}\nBroj ispravljenih stavki je: {count}\nPogledaj attachments za vise informacija." +
                    $"\nIspravljeni racuni se nalaze u (172.16.130.74) {ispravljeniRacuni}" + $"\n\n\n\nIzvrseno " + now.ToString("dd.MM.yyyy HH:mm:ss");
                //string att1 = dirLogFile + "\\" + now.ToString("dd_MM_yyyy") + "_" + "Log.txt";
                string att1 = dirLogFile + "\\" + now.ToString("dd_MM_yyyy") + "_" + "PromeneLog.txt";
                MailAddress to = new MailAddress("it@lilly.rs");

                SendEmail(to, body, att1);

                //waits for the mail to be send before closing the program
                Thread.Sleep(3000);
            }
            else
            {
                body += "Nije pronadjen ni jedan racun koji ima barkod prokucan u kolicini" +
                    $"\n\n\n\nIzvrseno " + now.ToString("dd.MM.yyyy HH:mm:ss");
                using (StreamWriter _log = new StreamWriter(dirLogFile + "\\" + now.ToString("dd_MM_yyyy") + "_" + "SveOkLog.txt", true))
                {
                    _log.WriteLine(now.ToString("G") + "\t"
                                + "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");

                    _log.WriteLine(now.ToString("G") + "\t"
                        + $"Nema racuna sa barkodom u kolicini");

                    _log.WriteLine(now.ToString("G") + "\t"
                                + "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!");
                }
                string att1 = dirLogFile + "\\" + now.ToString("dd_MM_yyyy") + "_" + "SveOkLog.txt";
                MailAddress to = new MailAddress("petar.jovancic@lilly.rs");

                SendEmail(to, body, att1);

                //waits for the mail to be send before closing the program
                Thread.Sleep(1000);
            }



        }//Main method

        //Metoda koja menja podatke, za pozitivnu kolicinu
        public static void LogIzmenaPozitivne(XElement postavka, string dirLogFile, DateTime now, PodaciRacuna infoRacun)
        {

            using (StreamWriter log = new StreamWriter(dirLogFile + "\\" + now.ToString("dd_MM_yyyy") + "_" + "PromeneLog.txt", true))
            {
                log.WriteLine(now.ToString("G") + "\t" + $"Broj racuna: {infoRacun.BrRacuna} | Datum i vreme: {infoRacun.DatumRacuna} {infoRacun.VremeRacuna} | " +
                    $"Poslovnica: {infoRacun.Poslovnica} | Blagajna: {infoRacun.Blagajna}");
                log.WriteLine(now.ToString("G") + "\t" + $"Pozicija: {postavka.Attribute("POZICIJA").Value}");
                log.WriteLine(now.ToString("G") + "\t" + $"Artikal: {postavka.Attribute("ARTIKEL").Value}");
                log.WriteLine(now.ToString("G") + "\t" + $"EAN: {postavka.Attribute("EAN").Value}");

                //postavlja kolicinu na 1
                log.WriteLine(now.ToString("G") + "\t" + $"Kolicina je bila: {postavka.Attribute("KOLICINA").Value}");
                postavka.Attribute("KOLICINA").Value = "1";
                log.WriteLine(now.ToString("G") + "\t" + $"Kolicina postavljena na: {postavka.Attribute("KOLICINA").Value}");

                //Postavlja vrednosti IZR_VREDNOST, NA_IZR_MPV, NA_IZR_MPV_BPOPDOK kao sto je NA_IZR_MPC
                log.WriteLine(now.ToString("G") + "\t" + $"IZR_VREDNOST je bila: {postavka.Attribute("IZR_VREDNOST").Value}");
                postavka.Attribute("IZR_VREDNOST").Value = postavka.Attribute("NA_IZR_MPC").Value;
                log.WriteLine(now.ToString("G") + "\t" + $"IZR_VREDNOST promenjena na: {postavka.Attribute("IZR_VREDNOST").Value}");
                log.WriteLine(now.ToString("G") + "\t" + $"NA_IZR_MPV  je bila: {postavka.Attribute("NA_IZR_MPV").Value}");
                postavka.Attribute("NA_IZR_MPV").Value = postavka.Attribute("NA_IZR_MPC").Value;
                log.WriteLine(now.ToString("G") + "\t" + $"NA_IZR_MPV promenjena na: {postavka.Attribute("NA_IZR_MPV").Value}");
                log.WriteLine(now.ToString("G") + "\t" + $"NA_IZR_MPV_BPOPDOK je bila: {postavka.Attribute("NA_IZR_MPV_BPOPDOK").Value}");
                postavka.Attribute("NA_IZR_MPV_BPOPDOK").Value = postavka.Attribute("NA_IZR_MPC").Value;
                log.WriteLine(now.ToString("G") + "\t" + $"NA_IZR_MPV_BPOPDOK promenjena na: {postavka.Attribute("NA_IZR_MPV_BPOPDOK").Value}");

                //postavlja vrednosti NA_IZR_PV i NA_IZR_PV_BPOPDOK na vrednost NA_IZR_PC
                log.WriteLine(now.ToString("G") + "\t" + $"NA_IZR_PV je bila: {postavka.Attribute("NA_IZR_PV").Value}");
                postavka.Attribute("NA_IZR_PV").Value = postavka.Attribute("NA_IZR_PC").Value;
                log.WriteLine(now.ToString("G") + "\t" + $"NA_IZR_PV promenjena na: {postavka.Attribute("NA_IZR_PV").Value}");
                log.WriteLine(now.ToString("G") + "\t" + $"NA_IZR_PV_BPOPDOK je bila: {postavka.Attribute("NA_IZR_PV_BPOPDOK").Value}");
                postavka.Attribute("NA_IZR_PV_BPOPDOK").Value = postavka.Attribute("NA_IZR_PC").Value;
                log.WriteLine(now.ToString("G") + "\t" + $"NA_IZR_PV_BPOPDOK promenjena na: {postavka.Attribute("NA_IZR_PV_BPOPDOK").Value}");

            }


        }

        //Metoda koja menja podatke, za negativnu kolicinu
        public static void LogIzmenaNegativne(XElement postavka, string dirLogFile, DateTime now)
        {
            using (StreamWriter log = new StreamWriter(dirLogFile + "\\" + now.ToString("dd_MM_yyyy") + "_" + "PromeneLog.txt", true))
            {
                log.WriteLine(now.ToString("G") + "\t" + $"Pozicija: {postavka.Attribute("POZICIJA").Value}");
                log.WriteLine(now.ToString("G") + "\t" + $"Artikal: {postavka.Attribute("ARTIKEL").Value}");
                log.WriteLine(now.ToString("G") + "\t" + $"EAN: {postavka.Attribute("EAN").Value}");

                //postavlja kolicinu na 1
                log.WriteLine(now.ToString("G") + "\t" + $"Kolicina je bila: {postavka.Attribute("KOLICINA").Value}");
                postavka.Attribute("KOLICINA").Value = "-1";
                log.WriteLine(now.ToString("G") + "\t" + $"Kolicina postavljena na: {postavka.Attribute("KOLICINA").Value}");

                //Postavlja vrednosti IZR_VREDNOST, NA_IZR_MPV, NA_IZR_MPV_BPOPDOK kao sto je NA_IZR_MPC
                log.WriteLine(now.ToString("G") + "\t" + $"IZR_VREDNOST je bila: {postavka.Attribute("IZR_VREDNOST").Value}");
                postavka.Attribute("IZR_VREDNOST").Value = "-" + postavka.Attribute("NA_IZR_MPC").Value;
                log.WriteLine(now.ToString("G") + "\t" + $"IZR_VREDNOST promenjena na: {postavka.Attribute("IZR_VREDNOST").Value}");
                log.WriteLine(now.ToString("G") + "\t" + $"NA_IZR_MPV  je bila: {postavka.Attribute("NA_IZR_MPV").Value}");
                postavka.Attribute("NA_IZR_MPV").Value = "-" + postavka.Attribute("NA_IZR_MPC").Value;
                log.WriteLine(now.ToString("G") + "\t" + $"NA_IZR_MPV promenjena na: {postavka.Attribute("NA_IZR_MPV").Value}");
                log.WriteLine(now.ToString("G") + "\t" + $"NA_IZR_MPV_BPOPDOK je bila: {postavka.Attribute("NA_IZR_MPV_BPOPDOK").Value}");
                postavka.Attribute("NA_IZR_MPV_BPOPDOK").Value = "-" + postavka.Attribute("NA_IZR_MPC").Value;
                log.WriteLine(now.ToString("G") + "\t" + $"NA_IZR_MPV_BPOPDOK promenjena na: {postavka.Attribute("NA_IZR_MPV_BPOPDOK").Value}");

                //postavlja vrednosti NA_IZR_PV i NA_IZR_PV_BPOPDOK na vrednost NA_IZR_PC
                log.WriteLine(now.ToString("G") + "\t" + $"NA_IZR_PV je bila: {postavka.Attribute("NA_IZR_PV").Value}");
                postavka.Attribute("NA_IZR_PV").Value = "-" + postavka.Attribute("NA_IZR_PC").Value;
                log.WriteLine(now.ToString("G") + "\t" + $"NA_IZR_PV promenjena na: {postavka.Attribute("NA_IZR_PV").Value}");
                log.WriteLine(now.ToString("G") + "\t" + $"NA_IZR_PV_BPOPDOK je bila: {postavka.Attribute("NA_IZR_PV_BPOPDOK").Value}");
                postavka.Attribute("NA_IZR_PV_BPOPDOK").Value = "-" + postavka.Attribute("NA_IZR_PC").Value;
                log.WriteLine(now.ToString("G") + "\t" + $"NA_IZR_PV_BPOPDOK promenjena na: {postavka.Attribute("NA_IZR_PV_BPOPDOK").Value}");



            }

        }



        //Method za slanje maila
        static void SendEmail(MailAddress to, string b, string att1)
        {
            MailAddress from = new MailAddress("EMAIL@lilly.rs");
            //MailAddress to = new MailAddress("petar.jovancic@lilly.rs");
            //MailAddress too = to;

            var smtpClinet = new SmtpClient("MAILSERVER")
            {
                Port = 587,
                Credentials = new NetworkCredential("EMAIL@lilly.rs", "PASSWORD"),
                EnableSsl = true,

            };


            using (MailMessage message = new MailMessage(from, to)
            {
                Subject = "Izvestaj KONTROLA",
                Body = b
            })
            {

                message.Attachments.Add(new Attachment(att1));
                //message.Attachments.Add(new Attachment(att2));
                smtpClinet.Send(message);

                Thread.Sleep(1000);

                smtpClinet.Dispose();

            }

        }


    }//Program class
}//namespace