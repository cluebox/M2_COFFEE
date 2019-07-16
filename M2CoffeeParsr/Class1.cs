using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using SpssLib.DataReader;
using SpssLib.SpssDataset;
using System.Data.SqlClient;

namespace M2CoffeeParsr
{
    class MayoCoffee
    {
        static void Main(string[] args)
        {
            int SurveyID = 600523;
            //string SurveyPeriod = "2014-09-30";//survey period
            string SurveyPeriod = "2019-04-30";
            string year = getYear(SurveyPeriod);
            string country = "INDONESIA";//survey country
            Insertion_m2coffee iobj = new Insertion_m2coffee();
            string questions = "iobs,weighting,r4,r13,r1,r6,r12,period,r21a,r21b_1,r21b_7,r21b_55,r21b_78,r21b_101,r21b_106,r21b_133,r21b_160,r21b_162,r21b_163,r21b_166,r21b_168,r21b_176,r24a_1,r24a_7,r24a_55,r24a_78,r24a_101,r24a_106,r24a_133,r24a_160,r24a_162,r24a_163,r24a_166,r24a_168,r24a_176,r21d,r21e_1,r21e_7,r21e_55,r21e_78,r21e_101,r21e_106,r21e_133,r21e_160,r21e_162,r21e_163,r21e_166,r21e_168,r21e_176,r24b_1,r24b_7,r24b_55,r24b_78,r24b_101,r24b_106,r24b_133,r24b_160,r24b_162,r24b_163,r24b_166,r24b_168,r24b_176,r21c,r25c,r123,r17_1,r17_2,r17_3,r17_4,r17_5,r17_6,r17_7,r17_8,r17_9,r17_10,r17_11,r17_12,r17_13,r17_14,r17_15,r18a_1,r18a_2,r18a_3,r18a_4,r18a_5,r18a_6,r18br1,r18br2,r18br3,r18br4,r18br5,r18br6,r15_1,r15_2,r15_3,r15_4,r15_5,r15_6,r15_7,r16,r25b_1,r25b_7,r25b_55,r25b_78,r25b_101,r25b_106,r25b_133,r25b_160,r25b_162,r25b_163,r25b_166,r25b_168,r25b_176,r25a_1,r25a_7,r25a_55,r25a_78,r25a_101,r25a_106,r25a_133,r25a_160,r25a_162,r25a_163,r25a_166,r25a_168,r25a_176,r29a,r29b,r124,r31a,r31b,r125,r33a,r33b,r126,r35a,r35b,r127,r37a,r37b,r128,r27h1_34,r27h1_35,r27h1_23,r27h1_10,r27h1_25,r27h1_30,r27h3_34,r27h3_35,r27h3_23,r27h3_10,r27h3_25,r27h3_30,r27h4_34,r27h4_35,r27h4_23,r27h4_10,r27h4_25,r27h4_30,r27h7_34,r27h7_35,r27h7_23,r27h7_10,r27h7_25,r27h7_30,r27h9_34,r27h9_35,r27h9_23,r27h9_10,r27h9_25,r27h9_30,r21anett,r21b_1,r21b_55,r21b_101,r21b_133,r21b_160,r24a_1,r24a_55,r24a_101,r24a_133,r24a_160,r21dnett,r21e_1,r21e_55,r21e_101,r21e_133,r21e_160,r24b_1,r24b_55,r24b_101,r24b_133,r24b_160,r25a_1,r25a_55,r25a_101,r25a_133,r25a_160,r25b_1,r25b_55,r25b_101,r25b_133,r25b_160,r21cnett,r25cnett,r123nett,r20_1,r20_2,r20_3,r20_4,r20_5,r20_6,r21b_266,r21b_267,r21b_268,r24a_266,r24a_267,r24a_268,r21e_266,r21e_267,r21e_268,r24b_266,r24b_267,r24b_268,r25b_266,r25b_267,r25b_268,r25a_266,r25a_267,r25a_268,r21b_281,r24a_281,r21e_281,r24b_281,r25a_281,r25b_281,r21b_64,r24a_64,r21e_64,r24b_64,r25a_64,r25b_64,r21b_272,r24a_272,r21e_272,r24b_272,r25b_272,r25a_272,r25d_1,r25d_7,r25d_55,r25d_78,r25d_101,r25d_106,r25d_133,r25d_160,r25d_162,r25d_163,r25d_166,r25d_168,r25d_176,r25d_266,r25d_267,r25d_268,r25d_272,r25d_64,r25e_1,r25e_7,r25e_55,r25e_78,r25e_101,r25e_106,r25e_133,r25e_160,r25e_162,r25e_163,r25e_166,r25e_168,r25e_176,r25e_266,r25e_267,r25e_268,r25e_272,r25e_64,r25e_281,r25d_281,r21b_300,r24a_300,r21e_300,r24b_300,r25b_300,r25a_300,r25e_300,r25d_300";// dashboard variable value
            //string questions = "iobs,weighting,period,r21a,r21b_64,r24a_64,r21d,r21e_64,r24b_64,r25a_64,r25b_64,r25c,r123,r21b_272,r24a_272,r21e_272,r24b_272,r25b_272,r25a_272,r25d_1 ,r25d_7 ,r25d_55 ,r25d_78 ,r25d_101 ,r25d_106 ,r25d_133 ,r25d_160 ,r25d_162 ,r25d_163 ,r25d_166 ,r25d_168 ,r25d_176 ,r25d_266 ,r25d_267 ,r25d_268 ,r25d_272 ,r25d_64 ,r25e_1 ,r25e_7 ,r25e_55 ,r25e_78 ,r25e_101 ,r25e_106 ,r25e_133 ,r25e_160 ,r25e_162 ,r25e_163 ,r25e_166 ,r25e_168 ,r25e_176 ,r25e_266 ,r25e_267 ,r25e_268 ,r25e_272 ,r25e_64 ,r25e_281 ,r25d_281";
            //string questions = "r25e_300,r25d_300";
            string[] spss_variable_name = questions.Split(',');
            SpssReader spssDataset;
            using (FileStream fileStream = new FileStream(@"D:\spssparsing\M2Coffee\M2Coffee-April19.sav", FileMode.Open, FileAccess.Read, FileShare.Read, 2048 * 10, FileOptions.SequentialScan))
            {
                spssDataset = new SpssReader(fileStream); // Create the reader, this will read the file header
                foreach (string spss_v in spss_variable_name)
                {
                    foreach (var variable in spssDataset.Variables)  // Iterate through all the varaibles
                    {
                        if (variable.Name.Equals(spss_v))
                        {
                            foreach (KeyValuePair<double, string> label in variable.ValueLabels)
                            {
                                string VARIABLE_NAME = spss_v;
                                string VARIABLE_NAME_SUB_NAME = variable.Name;
                                string VARIABLE_ID = label.Key.ToString();
                                string VARIABLE_VALUE = label.Value;
                                string VARIABLE_NAME_QUESTION = variable.Label;
                                string BASE_VARIABLE_NAME = variable.Name;
                                //iobj.insert_Dashboard_variable_values(VARIABLE_NAME, VARIABLE_NAME_SUB_NAME, VARIABLE_ID, VARIABLE_VALUE, VARIABLE_NAME_QUESTION, SurveyID, country, BASE_VARIABLE_NAME, SurveyPeriod);
                            }
                        }

                    }
                }
                foreach (var record in spssDataset.Records)
                {
                    string userID = null;
                    string variable_name;
                    string u_id = null;
                    decimal Weight = 0;
                    string Ses = "-- Not Available --";
                    string Period = "-- Not Available --";
                    string BrTom = "-- Not Available --";
                    string Gender = "-- Not Available --";
                    string MaritalStatus = "-- Not Available --";
                    string AttendedOn = "-- Not Available --";
                    string Location = "-- Not Available --";
                    string AgeGroup = "-- Not Available --";
                    string Yearly = "-- Not Available --";
                    string BrSpontABC = "-- Not Available --";
                    string BrSpontABCSusu = "-- Not Available --";
                    string BrSpontGoodDay = "-- Not Available --";
                    string BrSpontIndocafe = "-- Not Available --";
                    string BrSpontKapalApi = "-- Not Available --";
                    string BrSpontKASpecial = "-- Not Available --";
                    string BrSpontLuwak = "-- Not Available --";
                    string BrSpontTorabika = "-- Not Available --";
                    string BrSpontToraCappucino = "-- Not Available --";
                    string BrSpontToraDuo = "-- Not Available --";
                    string BrSpontToraMoka = "-- Not Available --";
                    string BrSpontToraSusu = "-- Not Available --";
                    string BrSpontTorabikaCreamyLatte = "-- Not Available --";
                    string BrAidABC = "-- Not Available --";
                    string BrAidABCSusu = "-- Not Available --";
                    string BrAidGoodDay = "-- Not Available --";
                    string BrAidIndocafe = "-- Not Available --";
                    string BrAidKapalApi = "-- Not Available --";
                    string BrAidKASpecial = "-- Not Available --";
                    string BrAidLuwak = "-- Not Available --";
                    string BrAidTorabika = "-- Not Available --";
                    string BrAidToraCappucino = "-- Not Available --";
                    string BrAidToraDuo = "-- Not Available --";
                    string BrAidToraMoka = "-- Not Available --";
                    string BrAidToraSusu = "-- Not Available --";
                    string BrAidTorabikaCreamyLatte = "-- Not Available --";
                    string AdTom = "-- Not Available --";
                    string AdSpontABC = "-- Not Available --";
                    string AdSpontABCSusu = "-- Not Available --";
                    string AdSpontGoodDay = "-- Not Available --";
                    string AdSpontIndocafe = "-- Not Available --";
                    string AdSpontKapalApi = "-- Not Available --";
                    string AdSpontKASpecial = "-- Not Available --";
                    string AdSpontLuwak = "-- Not Available --";
                    string AdSpontTorabika = "-- Not Available --";
                    string AdSpontToraCappucino = "-- Not Available --";
                    string AdSpontToraDuo = "-- Not Available --";
                    string AdSpontToraMoka = "-- Not Available --";
                    string AdSpontToraSusu = "-- Not Available --";
                    string AdSpontTorabikaCreamyLatte = "-- Not Available --";
                    string AdAidABC = "-- Not Available --";
                    string AdAidABCSusu = "-- Not Available --";
                    string AdAidGoodDay = "-- Not Available --";
                    string AdAidIndocafe = "-- Not Available --";
                    string AdAidKapalApi = "-- Not Available --";
                    string AdAidKASpecial = "-- Not Available --";
                    string AdAidLuwak = "-- Not Available --";
                    string AdAidTorabika = "-- Not Available --";
                    string AdAidToraCappucino = "-- Not Available --";
                    string AdAidToraDuo = "-- Not Available --";
                    string AdAidToraMoka = "-- Not Available --";
                    string AdAidToraSusu = "-- Not Available --";
                    string AdAidTorabikaCreamyLatte = "-- Not Available --";
                    string FavouriteBrand = "-- Not Available --";
                    string Bumo = "-- Not Available --";
                    string PrBumo = "-- Not Available --";
                    string RBCRTDTea = "-- Not Available --";
                    string RBCMineralwater = "-- Not Available --";
                    string RBCRTDJuice = "-- Not Available --";
                    string RBCCarbonatedRTD = "-- Not Available --";
                    string RBCRTDMilk = "-- Not Available --";
                    string RBCRTDCoffee = "-- Not Available --";
                    string RBCEnergydrink = "-- Not Available --";
                    string RBCVitamindrink = "-- Not Available --";
                    string RBCIsotonicdrink = "-- Not Available --";
                    string RBCYoghurtdrink = "-- Not Available --";
                    string RBCTeabag = "-- Not Available --";
                    string RBCInstantcoffee = "-- Not Available --";
                    string RBCInstantcoffeesachet = "-- Not Available --";
                    string RBCR_and_Gcoffeesachet = "-- Not Available --";
                    string RBCNone = "-- Not Available --";
                    string constypeP1mBlack = "-- Not Available --";
                    string constypeP1mmilk = "-- Not Available --";
                    string constypeP1mmocha = "-- Not Available --";
                    string constypeP1mWhite = "-- Not Available --";
                    string constypeP1mCappuccino = "-- Not Available --";
                    string constypeP1mRTD = "-- Not Available --";
                    string r18br1 = "-- Not Available --";
                    string r18br2 = "-- Not Available --";
                    string r18br3 = "-- Not Available --";
                    string r18br4 = "-- Not Available --";
                    string r18br5 = "-- Not Available --";
                    string r18br6 = "-- Not Available --";
                    string r15_1 = "-- Not Available --";
                    string r15_2 = "-- Not Available --";
                    string r15_3 = "-- Not Available --";
                    string r15_4 = "-- Not Available --";
                    string r15_5 = "-- Not Available --";
                    string r15_6 = "-- Not Available --";
                    string r15_7 = "-- Not Available --";
                    string DURATIONWATCHINGTV = "-- Not Available --";
                    string CurconABC = "-- Not Available --";
                    string CurconABCSusu = "-- Not Available --";
                    string CurconGoodDay = "-- Not Available --";
                    string CurconIndocafe = "-- Not Available --";
                    string CurconKapalApi = "-- Not Available --";
                    string CurconKASpecial = "-- Not Available --";
                    string CurconLuwak = "-- Not Available --";
                    string CurconTorabika = "-- Not Available --";
                    string CurconToraCappucino = "-- Not Available --";
                    string CurconToraDuo = "-- Not Available --";
                    string CurconToraMoka = "-- Not Available --";
                    string CurconToraSusu = "-- Not Available --";
                    string CurconTorabikaCreamyLatte = "-- Not Available --";
                    string ConLMABC = "-- Not Available --";
                    string ConLMABCSusu = "-- Not Available --";
                    string ConLMGoodDay = "-- Not Available --";
                    string ConLMIndocafe = "-- Not Available --";
                    string ConLMKapalApi = "-- Not Available --";
                    string ConLMKASpecial = "-- Not Available --";
                    string ConLMLuwak = "-- Not Available --";
                    string ConLMTorabika = "-- Not Available --";
                    string ConLMToraCappucino = "-- Not Available --";
                    string ConLMToraDuo = "-- Not Available --";
                    string ConLMToraMoka = "-- Not Available --";
                    string ConLMToraSusu = "-- Not Available --";
                    string ConLMTorabikaCreamyLatte = "-- Not Available --";
                    string Black_Coffee_Fav_Br = "-- Not Available --";
                    string Black_Coffee_Bumo = "-- Not Available --";
                    string Black_Coffee_PreBumo = "-- Not Available --";
                    string Milk_Coffee_Fav_Br = "-- Not Available --";
                    string Milk_Coffee_Bumo = "-- Not Available --";
                    string Milk_Coffee_PreBumo = "-- Not Available --";
                    string Mocca_Coffee_Fav_Br = "-- Not Available --";
                    string Mocca_Coffee_Bumo = "-- Not Available --";
                    string Mocca_Coffee_PreBumo = "-- Not Available --";
                    string White_Coffee_Fav_Br = "-- Not Available --";
                    string White_Coffee_Bumo = "-- Not Available --";
                    string White_Coffee_PreBumo = "-- Not Available --";
                    string Cappucinno_Fav_Br = "-- Not Available --";
                    string Cappucinno_Bumo = "-- Not Available --";
                    string Cappucinno_PreBumo = "-- Not Available --";
                    string r27h1_1 = "-- Not Available --";
                    string r27h1_22 = "-- Not Available --";
                    string r27h1_26 = "-- Not Available --";
                    string r27h1_10 = "-- Not Available --";
                    string r27h1_25 = "-- Not Available --";
                    string r27h1_30 = "-- Not Available --";
                    string r27h3_1 = "-- Not Available --";
                    string r27h3_22 = "-- Not Available --";
                    string r27h3_26 = "-- Not Available --";
                    string r27h3_10 = "-- Not Available --";
                    string r27h3_25 = "-- Not Available --";
                    string r27h3_30 = "-- Not Available --";
                    string r27h4_1 = "-- Not Available --";
                    string r27h4_22 = "-- Not Available --";
                    string r27h4_26 = "-- Not Available --";
                    string r27h4_10 = "-- Not Available --";
                    string r27h4_25 = "-- Not Available --";
                    string r27h4_30 = "-- Not Available --";
                    string r27h7_1 = "-- Not Available --";
                    string r27h7_22 = "-- Not Available --";
                    string r27h7_26 = "-- Not Available --";
                    string r27h7_10 = "-- Not Available --";
                    string r27h7_25 = "-- Not Available --";
                    string r27h7_30 = "-- Not Available --";
                    string r27h9_1 = "-- Not Available --";
                    string r27h9_22 = "-- Not Available --";
                    string r27h9_26 = "-- Not Available --";
                    string r27h9_10 = "-- Not Available --";
                    string r27h9_25 = "-- Not Available --";
                    string r27h9_30 = "-- Not Available --";
                    string r21b_1 = "-- Not Available --";
                    string r21b_55 = "-- Not Available --";
                    string r21b_101 = "-- Not Available --";
                    string r21b_133 = "-- Not Available --";
                    string r21b_160 = "-- Not Available --";
                    string r24a_1 = "-- Not Available --";
                    string r24a_55 = "-- Not Available --";
                    string r24a_101 = "-- Not Available --";
                    string r24a_133 = "-- Not Available --";
                    string r24a_160 = "-- Not Available --";
                    string r21e_1 = "-- Not Available --";
                    string r21e_55 = "-- Not Available --";
                    string r21e_101 = "-- Not Available --";
                    string r21e_133 = "-- Not Available --";
                    string r21e_160 = "-- Not Available --";
                    string r24b_1 = "-- Not Available --";
                    string r24b_55 = "-- Not Available --";
                    string r24b_101 = "-- Not Available --";
                    string r24b_133 = "-- Not Available --";
                    string r24b_160 = "-- Not Available --";
                    string r25a_1 = "-- Not Available --";
                    string r25a_55 = "-- Not Available --";
                    string r25a_101 = "-- Not Available --";
                    string r25a_133 = "-- Not Available --";
                    string r25a_160 = "-- Not Available --";
                    string r25b_1 = "-- Not Available --";
                    string r25b_55 = "-- Not Available --";
                    string r25b_101 = "-- Not Available --";
                    string r25b_133 = "-- Not Available --";
                    string r25b_160 = "-- Not Available --";
                    string r27h1_34 = "-- Not Available --";
                    string r27h1_35 = "-- Not Available --";
                    string r27h1_23 = "-- Not Available --";
                    string r27h3_34 = "-- Not Available --";
                    string r27h3_35 = "-- Not Available --";
                    string r27h3_23 = "-- Not Available --";
                    string r27h4_34 = "-- Not Available --";
                    string r27h4_35 = "-- Not Available --";
                    string r27h4_23 = "-- Not Available --";
                    string r27h7_34 = "-- Not Available --";
                    string r27h7_35 = "-- Not Available --";
                    string r27h7_23 = "-- Not Available --";
                    string r27h9_34 = "-- Not Available --";
                    string r27h9_35 = "-- Not Available --";
                    string r27h9_23 = "-- Not Available --";
                    string r21anett = "-- Not Available --";
                    string r21dnett = "-- Not Available --";
                    string r21cnett = "-- Not Available --";
                    string r25cnett = "-- Not Available --";
                    string r123nett = "-- Not Available --";
                    string BrTomNet = "-- Not Available --";
                    string AdTomNet = "-- Not Available --";
                    string Net_Fav = "-- Not Available --";
                    string Net_Bumo = "-- Not Available --";
                    string Net_PreBumo = "-- Not Available --";
                    string r20_1 = "-- Not Available --";
                    string r20_2 = "-- Not Available --";
                    string r20_3 = "-- Not Available --";
                    string r20_4 = "-- Not Available --";
                    string r20_5 = "-- Not Available --";
                    string r20_6 = "-- Not Available --";
                    string BrSpontToraCafe_Volcano_Chocomelt = "-- Not Available --";
                    string BrSpontTora_Cafe_Caramelove = "-- Not Available --";
                    string BrSpontTora_Cafe_tidak_spesifik = "-- Not Available --";
                    string BrAidToraCafe_Volcano_Chocomelt = "-- Not Available --";
                    string BrAidTora_Cafe_Caramelove = "-- Not Available --";
                    string BrAidTora_Cafe_tidak_spesifik = "-- Not Available --";
                    string AdSpontToraCafe_Volcano_Chocomelt = "-- Not Available --";
                    string AdSpontTora_Cafe_Caramelove = "-- Not Available --";
                    string AdSpontTora_Cafe_tidak_spesifik = "-- Not Available --";
                    string AdAidToraCafe_Volcano_Chocomelt = "-- Not Available --";
                    string AdAidTora_Cafe_Caramelove = "-- Not Available --";
                    string AdAidTora_Cafe_tidak_spesifik = "-- Not Available --";
                    string CurconToraCafe_Volcano_Chocomelt = "-- Not Available --";
                    string CurconTora_Cafe_Caramelove = "-- Not Available --";
                    string CurconTora_Cafe_tidak_spesifik = "-- Not Available --";
                    string ConL3MToraCafe_Volcano_Chocomelt = "-- Not Available --";
                    string ConL3MTora_Cafe_Caramelove = "-- Not Available --";
                    string ConL3MTora_Cafe_tidak_spesifik = "-- Not Available --";
                    string BrSpontToraCafe = "-- Not Available --";
                    string BrAidToraCafe = "-- Not Available --";
                    string AdSpontToraCafe = "-- Not Available --";
                    string AdAidToraCafe = "-- Not Available --";
                    string ConsL3MToraCafe = "-- Not Available --";
                    string CurconToraCafe = "-- Not Available --";
                    string BrSpont_GoodDayCapucino = "-- Not Available --";
                    string BrAid_GoodDayCapucino = "-- Not Available --";
                    string AdSpont_GoodDayCapucino = "-- Not Available --";
                    string AdAid_GoodDayCapucino = "-- Not Available --";
                    string Curcon_GoodDayCapucino = "-- Not Available --";
                    string ConsL3M_GoodDayCapucino = "-- Not Available --";
                    string BrSpontTora_Kopi_Susu_Esp = "-- Not Available --";
                    string BrAid_Tora_Kopi_Susu_Esp = "-- Not Available --";
                    string AdSpont_Tora_Kopi_Susu_Esp = "-- Not Available --";
                    string AdAid_Tora_Kopi_Susu_Esp = "-- Not Available --";
                    string Curcon_Tora_Kopi_Susu_Esp = "-- Not Available --";
                    string ConsL3M_Tora_Kopi_Susu_Esp = "-- Not Available --";
                    string ConP1WABC = "-- Not Available --";
                    string ConP1WABCSusu = "-- Not Available --";
                    string ConP1WGoodDay = "-- Not Available --";
                    string ConP1WIndocafe = "-- Not Available --";
                    string ConP1WKapalApi = "-- Not Available --";
                    string ConP1WKASpecial = "-- Not Available --";
                    string ConP1WLuwak = "-- Not Available --";
                    string ConP1WTorabika = "-- Not Available --";
                    string ConP1WToraCappucino = "-- Not Available --";
                    string ConP1WToraDuo = "-- Not Available --";
                    string ConP1WToraMoka = "-- Not Available --";
                    string ConP1WToraSusu = "-- Not Available --";
                    string ConP1WTorabikaCreamyLatte = "-- Not Available --";
                    string ConP1WToraCafe_Volcano_Chocomelt = "-- Not Available --";
                    string ConP1WTora_Cafe_Caramelove = "-- Not Available --";
                    string ConP1WTora_Cafe_tidak_spesifik = "-- Not Available --";
                    string ConsP1W_Tora_Kopi_Susu_Esp = "-- Not Available --";
                    string ConsP1W_GoodDayCapucino = "-- Not Available --";
                    string ConP1DABC = "-- Not Available --";
                    string ConP1DABCSusu = "-- Not Available --";
                    string ConP1DGoodDay = "-- Not Available --";
                    string ConP1DIndocafe = "-- Not Available --";
                    string ConP1DKapalApi = "-- Not Available --";
                    string ConP1DKASpecial = "-- Not Available --";
                    string ConP1DLuwak = "-- Not Available --";
                    string ConP1DTorabika = "-- Not Available --";
                    string ConP1DToraCappucino = "-- Not Available --";
                    string ConP1DToraDuo = "-- Not Available --";
                    string ConP1DToraMoka = "-- Not Available --";
                    string ConP1DToraSusu = "-- Not Available --";
                    string ConP1DTorabikaCreamyLatte = "-- Not Available --";
                    string ConP1DToraCafe_Volcano_Chocomelt = "-- Not Available --";
                    string ConP1DTora_Cafe_Caramelove = "-- Not Available --";
                    string ConP1DTora_Cafe_tidak_spesifik = "-- Not Available --";
                    string ConsP1D_Tora_Kopi_Susu_Esp = "-- Not Available --";
                    string ConsP1D_GoodDayCapucino = "-- Not Available --";
                    string ConsP1DToraCafe = "-- Not Available --";
                    string ConsP1WToraCafe = "-- Not Available --";
                    string BrSpontTOP_Coffee_Cappuccino = "-- Not Available --";
                    string BrAid_TOP_Coffee_Cappuccino = "-- Not Available --";
                    string AdSpont_TOP_Coffee_Cappuccino = "-- Not Available --";
                    string AdAid_TOP_Coffee_Cappuccino = "-- Not Available --";
                    string Curcon_TOP_Coffee_Cappuccino = "-- Not Available --";
                    string ConsL3M_TOP_Coffee_Cappuccino = "-- Not Available --";
                    string ConsP1DTOP_Coffee_Cappuccino = "-- Not Available --";
                    string ConsP1wTOP_Coffee_Cappuccino = "-- Not Available --";

                    foreach (var variable in spssDataset.Variables)
                    {
                        foreach (string spss_v in spss_variable_name)
                        {
                            if (variable.Name.Equals(spss_v))
                            {
                                variable_name = variable.Name;
                                switch (variable_name)
                                {
                                    case "iobs":
                                        {
                                            u_id = Convert.ToString(record.GetValue(variable));
                                            userID = find_UserId(SurveyID, SurveyPeriod, u_id);
                                            //userID = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "weighting":
                                        {
                                            Weight = Convert.ToDecimal(record.GetValue(variable));
                                            break;
                                        }
                                    case "r4":
                                        {
                                            Gender = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r13":
                                        {
                                            MaritalStatus = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r1":
                                        {
                                            Location = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r6":
                                        {
                                            AgeGroup = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r12":
                                        {
                                            Ses = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "period":
                                        {
                                            Period = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21a":
                                        {
                                            BrTom = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21b_1":
                                        {
                                            BrSpontABC = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21b_7":
                                        {
                                            BrSpontABCSusu = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21b_55":
                                        {
                                            BrSpontGoodDay = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21b_78":
                                        {
                                            BrSpontIndocafe = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21b_101":
                                        {
                                            BrSpontKapalApi = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21b_106":
                                        {
                                            BrSpontKASpecial = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21b_133":
                                        {
                                            BrSpontLuwak = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21b_160":
                                        {
                                            BrSpontTorabika = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21b_162":
                                        {
                                            BrSpontToraCappucino = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21b_163":
                                        {
                                            BrSpontToraDuo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21b_166":
                                        {
                                            BrSpontToraMoka = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21b_168":
                                        {
                                            BrSpontToraSusu = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21b_176":
                                        {
                                            BrSpontTorabikaCreamyLatte = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24a_1":
                                        {
                                            BrAidABC = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24a_7":
                                        {
                                            BrAidABCSusu = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24a_55":
                                        {
                                            BrAidGoodDay = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24a_78":
                                        {
                                            BrAidIndocafe = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24a_101":
                                        {
                                            BrAidKapalApi = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24a_106":
                                        {
                                            BrAidKASpecial = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24a_133":
                                        {
                                            BrAidLuwak = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24a_160":
                                        {
                                            BrAidTorabika = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24a_162":
                                        {
                                            BrAidToraCappucino = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24a_163":
                                        {
                                            BrAidToraDuo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24a_166":
                                        {
                                            BrAidToraMoka = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24a_168":
                                        {
                                            BrAidToraSusu = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24a_176":
                                        {
                                            BrAidTorabikaCreamyLatte = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21d":
                                        {
                                            AdTom = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21e_1":
                                        {
                                            AdSpontABC = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21e_7":
                                        {
                                            AdSpontABCSusu = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21e_55":
                                        {
                                            AdSpontGoodDay = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21e_78":
                                        {
                                            AdSpontIndocafe = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21e_101":
                                        {
                                            AdSpontKapalApi = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21e_106":
                                        {
                                            AdSpontKASpecial = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21e_133":
                                        {
                                            AdSpontLuwak = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21e_160":
                                        {
                                            AdSpontTorabika = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21e_162":
                                        {
                                            AdSpontToraCappucino = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21e_163":
                                        {
                                            AdSpontToraDuo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21e_166":
                                        {
                                            AdSpontToraMoka = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21e_168":
                                        {
                                            AdSpontToraSusu = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21e_176":
                                        {
                                            AdSpontTorabikaCreamyLatte = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24b_1":
                                        {
                                            AdAidABC = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24b_7":
                                        {
                                            AdAidABCSusu = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24b_55":
                                        {
                                            AdAidGoodDay = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24b_78":
                                        {
                                            AdAidIndocafe = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24b_101":
                                        {
                                            AdAidKapalApi = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24b_106":
                                        {
                                            AdAidKASpecial = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24b_133":
                                        {
                                            AdAidLuwak = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24b_160":
                                        {
                                            AdAidTorabika = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24b_162":
                                        {
                                            AdAidToraCappucino = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24b_163":
                                        {
                                            AdAidToraDuo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24b_166":
                                        {
                                            AdAidToraMoka = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24b_168":
                                        {
                                            AdAidToraSusu = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24b_176":
                                        {
                                            AdAidTorabikaCreamyLatte = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21c":
                                        {
                                            FavouriteBrand = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25c":
                                        {
                                            Bumo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r123":
                                        {
                                            PrBumo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r17_1":
                                        {
                                            RBCRTDTea = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r17_2":
                                        {
                                            RBCMineralwater = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r17_3":
                                        {
                                            RBCRTDJuice = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r17_4":
                                        {
                                            RBCCarbonatedRTD = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r17_5":
                                        {
                                            RBCRTDMilk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r17_6":
                                        {
                                            RBCRTDCoffee = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r17_7":
                                        {
                                            RBCEnergydrink = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r17_8":
                                        {
                                            RBCVitamindrink = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r17_9":
                                        {
                                            RBCIsotonicdrink = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r17_10":
                                        {
                                            RBCYoghurtdrink = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r17_11":
                                        {
                                            RBCTeabag = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r17_12":
                                        {
                                            RBCInstantcoffee = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r17_13":
                                        {
                                            RBCInstantcoffeesachet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r17_14":
                                        {
                                            RBCR_and_Gcoffeesachet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r17_15":
                                        {
                                            RBCNone = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r18a_1":
                                        {
                                            constypeP1mBlack = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r18a_2":
                                        {
                                            constypeP1mmilk = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r18a_3":
                                        {
                                            constypeP1mmocha = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r18a_4":
                                        {
                                            constypeP1mWhite = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r18a_5":
                                        {
                                            constypeP1mCappuccino = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r18a_6":
                                        {
                                            constypeP1mRTD = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r18br1":
                                        {
                                            r18br1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r18br2":
                                        {
                                            r18br2 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r18br3":
                                        {
                                            r18br3 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r18br4":
                                        {
                                            r18br4 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r18br5":
                                        {
                                            r18br5 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r18br6":
                                        {
                                            r18br6 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r15_1":
                                        {
                                            r15_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r15_2":
                                        {
                                            r15_2 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r15_3":
                                        {
                                            r15_3 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r15_4":
                                        {
                                            r15_4 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r15_5":
                                        {
                                            r15_5 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r15_6":
                                        {
                                            r15_6 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r15_7":
                                        {
                                            r15_7 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r16":
                                        {
                                            DURATIONWATCHINGTV = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25b_1":
                                        {
                                            CurconABC = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25b_7":
                                        {
                                            CurconABCSusu = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25b_55":
                                        {
                                            CurconGoodDay = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25b_78":
                                        {
                                            CurconIndocafe = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25b_101":
                                        {
                                            CurconKapalApi = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25b_106":
                                        {
                                            CurconKASpecial = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25b_133":
                                        {
                                            CurconLuwak = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25b_160":
                                        {
                                            CurconTorabika = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25b_162":
                                        {
                                            CurconToraCappucino = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25b_163":
                                        {
                                            CurconToraDuo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25b_166":
                                        {
                                            CurconToraMoka = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25b_168":
                                        {
                                            CurconToraSusu = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25b_176":
                                        {
                                            CurconTorabikaCreamyLatte = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25a_1":
                                        {
                                            ConLMABC = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25a_7":
                                        {
                                            ConLMABCSusu = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25a_55":
                                        {
                                            ConLMGoodDay = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25a_78":
                                        {
                                            ConLMIndocafe = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25a_101":
                                        {
                                            ConLMKapalApi = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25a_106":
                                        {
                                            ConLMKASpecial = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25a_133":
                                        {
                                            ConLMLuwak = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25a_160":
                                        {
                                            ConLMTorabika = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25a_162":
                                        {
                                            ConLMToraCappucino = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25a_163":
                                        {
                                            ConLMToraDuo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25a_166":
                                        {
                                            ConLMToraMoka = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25a_168":
                                        {
                                            ConLMToraSusu = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25a_176":
                                        {
                                            ConLMTorabikaCreamyLatte = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r29a":
                                        {
                                            Black_Coffee_Fav_Br = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r29b":
                                        {
                                            Black_Coffee_Bumo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r124":
                                        {
                                            Black_Coffee_PreBumo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r31a":
                                        {
                                            Milk_Coffee_Fav_Br = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r31b":
                                        {
                                            Milk_Coffee_Bumo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r125":
                                        {
                                            Milk_Coffee_PreBumo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r33a":
                                        {
                                            Mocca_Coffee_Fav_Br = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r33b":
                                        {
                                            Mocca_Coffee_Bumo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r126":
                                        {
                                            Mocca_Coffee_PreBumo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r35a":
                                        {
                                            White_Coffee_Fav_Br = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r35b":
                                        {
                                            White_Coffee_Bumo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r127":
                                        {
                                            White_Coffee_PreBumo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r37a":
                                        {
                                            Cappucinno_Fav_Br = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r37b":
                                        {
                                            Cappucinno_Bumo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r128":
                                        {
                                            Cappucinno_PreBumo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h1_1":
                                        {
                                            r27h1_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h1_22":
                                        {
                                            r27h1_22 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h1_26":
                                        {
                                            r27h1_26 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h1_10":
                                        {
                                            r27h1_10 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h1_25":
                                        {
                                            r27h1_25 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h1_30":
                                        {
                                            r27h1_30 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h3_1":
                                        {
                                            r27h3_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h3_22":
                                        {
                                            r27h3_22 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h3_26":
                                        {
                                            r27h3_26 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h3_10":
                                        {
                                            r27h3_10 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h3_25":
                                        {
                                            r27h3_25 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h3_30":
                                        {
                                            r27h3_30 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h4_1":
                                        {
                                            r27h4_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h4_22":
                                        {
                                            r27h4_22 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h4_26":
                                        {
                                            r27h4_26 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h4_10":
                                        {
                                            r27h4_10 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h4_25":
                                        {
                                            r27h4_25 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h4_30":
                                        {
                                            r27h4_30 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h7_1":
                                        {
                                            r27h7_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h7_22":
                                        {
                                            r27h7_22 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h7_26":
                                        {
                                            r27h7_26 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h7_10":
                                        {
                                            r27h7_10 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h7_25":
                                        {
                                            r27h7_25 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h7_30":
                                        {
                                            r27h7_30 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h9_1":
                                        {
                                            r27h9_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h9_22":
                                        {
                                            r27h9_22 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h9_26":
                                        {
                                            r27h9_26 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h9_10":
                                        {
                                            r27h9_10 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h9_25":
                                        {
                                            r27h9_25 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h9_30":
                                        {
                                            r27h9_30 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r27h1_34":
                                        {
                                            r27h1_34 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h1_35":
                                        {
                                            r27h1_35 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h1_23":
                                        {
                                            r27h1_23 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h3_34":
                                        {
                                            r27h3_34 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h3_35":
                                        {
                                            r27h3_35 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h3_23":
                                        {
                                            r27h3_23 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h4_34":
                                        {
                                            r27h4_34 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h4_35":
                                        {
                                            r27h4_35 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h4_23":
                                        {
                                            r27h4_23 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h7_34":
                                        {
                                            r27h7_34 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h7_35":
                                        {
                                            r27h7_35 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h7_23":
                                        {
                                            r27h7_23 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h9_34":
                                        {
                                            r27h9_34 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h9_35":
                                        {
                                            r27h9_35 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r27h9_23":
                                        {
                                            r27h9_23 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21anett":
                                        {
                                            BrTomNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21dnett":
                                        {
                                            AdTomNet = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21cnett":
                                        {
                                            Net_Fav = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25cnett":
                                        {
                                            Net_Bumo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r123nett":
                                        {
                                            Net_PreBumo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r20_1":
                                        {
                                            r20_1 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r20_2":
                                        {
                                            r20_2 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r20_3":
                                        {
                                            r20_3 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r20_4":
                                        {
                                            r20_4 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r20_5":
                                        {
                                            r20_5 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r20_6":
                                        {
                                            r20_6 = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21b_266":
                                        {
                                            BrSpontToraCafe_Volcano_Chocomelt = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21b_267":
                                        {
                                            BrSpontTora_Cafe_Caramelove = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21b_268":
                                        {
                                            BrSpontTora_Cafe_tidak_spesifik = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24a_266":
                                        {
                                            BrAidToraCafe_Volcano_Chocomelt = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24a_267":
                                        {
                                            BrAidTora_Cafe_Caramelove = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24a_268":
                                        {
                                            BrAidTora_Cafe_tidak_spesifik = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21e_266":
                                        {
                                            AdSpontToraCafe_Volcano_Chocomelt = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21e_267":
                                        {
                                            AdSpontTora_Cafe_Caramelove = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21e_268":
                                        {
                                            AdSpontTora_Cafe_tidak_spesifik = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24b_266":
                                        {
                                            AdAidToraCafe_Volcano_Chocomelt = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24b_267":
                                        {
                                            AdAidTora_Cafe_Caramelove = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24b_268":
                                        {
                                            AdAidTora_Cafe_tidak_spesifik = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25b_266":
                                        {
                                            CurconToraCafe_Volcano_Chocomelt = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25b_267":
                                        {
                                            CurconTora_Cafe_Caramelove = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25b_268":
                                        {
                                            CurconTora_Cafe_tidak_spesifik = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25a_266":
                                        {
                                            ConL3MToraCafe_Volcano_Chocomelt = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25a_267":
                                        {
                                            ConL3MTora_Cafe_Caramelove = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25a_268":
                                        {
                                            ConL3MTora_Cafe_tidak_spesifik = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21b_281":
                                        {
                                            BrSpontToraCafe = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24a_281":
                                        {
                                            BrAidToraCafe = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21e_281":
                                        {
                                            AdSpontToraCafe = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24b_281":
                                        {
                                            AdAidToraCafe = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25a_281":
                                        {
                                            ConsL3MToraCafe = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25b_281":
                                        {
                                            CurconToraCafe = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21b_64":
                                        {
                                            BrSpont_GoodDayCapucino = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24a_64":
                                        {
                                            BrAid_GoodDayCapucino = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21e_64":
                                        {
                                            AdSpont_GoodDayCapucino = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24b_64":
                                        {
                                            AdAid_GoodDayCapucino = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25a_64":
                                        {
                                            ConsL3M_GoodDayCapucino = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25b_64":
                                        {
                                            Curcon_GoodDayCapucino = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21b_272":
                                        {
                                            BrSpontTora_Kopi_Susu_Esp = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24a_272":
                                        {
                                            BrAid_Tora_Kopi_Susu_Esp = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r21e_272":
                                        {
                                            AdSpont_Tora_Kopi_Susu_Esp = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r24b_272":
                                        {
                                            AdAid_Tora_Kopi_Susu_Esp = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25b_272":
                                        {
                                            Curcon_Tora_Kopi_Susu_Esp = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25a_272":
                                        {
                                            ConsL3M_Tora_Kopi_Susu_Esp = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25d_1":
                                        {
                                            ConP1WABC = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25d_7":
                                        {
                                            ConP1WABCSusu = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25d_55":
                                        {
                                            ConP1WGoodDay = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25d_78":
                                        {
                                            ConP1WIndocafe = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25d_101":
                                        {
                                            ConP1WKapalApi = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25d_106":
                                        {
                                            ConP1WKASpecial = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25d_133":
                                        {
                                            ConP1WLuwak = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25d_160":
                                        {
                                            ConP1WTorabika = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25d_162":
                                        {
                                            ConP1WToraCappucino = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25d_163":
                                        {
                                            ConP1WToraDuo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25d_166":
                                        {
                                            ConP1WToraMoka = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25d_168":
                                        {
                                            ConP1WToraSusu = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25d_176":
                                        {
                                            ConP1WTorabikaCreamyLatte = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25d_266":
                                        {
                                            ConP1WToraCafe_Volcano_Chocomelt = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25d_267":
                                        {
                                            ConP1WTora_Cafe_Caramelove = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25d_268":
                                        {
                                            ConP1WTora_Cafe_tidak_spesifik = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25d_272":
                                        {
                                            ConsP1W_Tora_Kopi_Susu_Esp = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r25d_64":
                                        {
                                            ConsP1W_GoodDayCapucino = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25e_1":
                                        {
                                            ConP1DABC = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25e_7":
                                        {
                                            ConP1DABCSusu = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25e_55":
                                        {
                                            ConP1DGoodDay = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25e_78":
                                        {
                                            ConP1DIndocafe = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25e_101":
                                        {
                                            ConP1DKapalApi = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25e_106":
                                        {
                                            ConP1DKASpecial = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25e_133":
                                        {
                                            ConP1DLuwak = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25e_160":
                                        {
                                            ConP1DTorabika = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25e_162":
                                        {
                                            ConP1DToraCappucino = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25e_163":
                                        {
                                            ConP1DToraDuo = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25e_166":
                                        {
                                            ConP1DToraMoka = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25e_168":
                                        {
                                            ConP1DToraSusu = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25e_176":
                                        {
                                            ConP1DTorabikaCreamyLatte = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25e_266":
                                        {
                                            ConP1DToraCafe_Volcano_Chocomelt = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25e_267":
                                        {
                                            ConP1DTora_Cafe_Caramelove = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25e_268":
                                        {
                                            ConP1DTora_Cafe_tidak_spesifik = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25e_272":
                                        {
                                            ConsP1D_Tora_Kopi_Susu_Esp = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }

                                    case "r25e_64":
                                        {
                                            ConsP1D_GoodDayCapucino = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25e_281":
                                        {
                                            ConsP1DToraCafe = Convert.ToString(record.GetValue(variable));
                                            break;
                                        }
                                    case "r25d_281":
                                        {
                                            ConsP1WToraCafe = Convert.ToString(record.GetValue(variable));

                                            break;
                                        }
                                    case "r21b_300": { BrSpontTOP_Coffee_Cappuccino = Convert.ToString(record.GetValue(variable)); break; }
                                    case "r24a_300": { BrAid_TOP_Coffee_Cappuccino = Convert.ToString(record.GetValue(variable)); break; }
                                    case "r21e_300": { AdSpont_TOP_Coffee_Cappuccino = Convert.ToString(record.GetValue(variable)); break; }
                                    case "r24b_300": { AdAid_TOP_Coffee_Cappuccino = Convert.ToString(record.GetValue(variable)); break; }
                                    case "r25b_300": { Curcon_TOP_Coffee_Cappuccino = Convert.ToString(record.GetValue(variable)); break; }
                                    case "r25a_300": { ConsL3M_TOP_Coffee_Cappuccino = Convert.ToString(record.GetValue(variable)); break; }
                                    case "r25e_300": { string ConsP1DTOP_Coffee_Cappuccino = Convert.ToString(record.GetValue(variable)); break; }
                                    case "r25d_300": { string ConsP1wTOP_Coffee_Cappuccino = Convert.ToString(record.GetValue(variable)); break; }
                                }
                            }
                        }
                    }
                    if (userID != null && Weight != 0)
                    {
                        iobj.insert_Dashboard_values(userID, Gender, MaritalStatus, country, SurveyID, SurveyPeriod, Weight, Location, AgeGroup, Ses, Period, year, BrTom, BrSpontABC, BrSpontGoodDay, BrSpontABCSusu, BrSpontIndocafe, BrSpontKapalApi, BrSpontKASpecial, BrSpontLuwak, BrSpontTorabika, BrSpontToraCappucino, BrSpontToraDuo, BrSpontToraMoka, BrSpontToraSusu, BrSpontTorabikaCreamyLatte, BrAidABC, BrAidABCSusu, BrAidGoodDay, BrAidIndocafe, BrAidKapalApi, BrAidKASpecial, BrAidLuwak, BrAidTorabika, BrAidToraCappucino, BrAidToraDuo, BrAidToraMoka, BrAidToraSusu, BrAidTorabikaCreamyLatte, AdTom, AdSpontABC, AdSpontABCSusu, AdSpontGoodDay, AdSpontIndocafe, AdSpontKapalApi, AdSpontKASpecial, AdSpontLuwak, AdSpontTorabika, AdSpontToraCappucino, AdSpontToraDuo, AdSpontToraMoka, AdSpontToraSusu, AdSpontTorabikaCreamyLatte, AdAidABC, AdAidABCSusu, AdAidGoodDay, AdAidIndocafe, AdAidKapalApi, AdAidKASpecial, AdAidLuwak, AdAidTorabika, AdAidToraCappucino, AdAidToraDuo, AdAidToraMoka, AdAidToraSusu, AdAidTorabikaCreamyLatte, FavouriteBrand, Bumo, PrBumo, RBCRTDTea, RBCMineralwater, RBCRTDJuice, RBCCarbonatedRTD, RBCRTDMilk, RBCRTDCoffee, RBCEnergydrink, RBCVitamindrink, RBCIsotonicdrink, RBCYoghurtdrink, RBCTeabag, RBCInstantcoffee, RBCInstantcoffeesachet, RBCR_and_Gcoffeesachet, RBCNone, constypeP1mBlack, constypeP1mmilk, constypeP1mmocha, constypeP1mWhite, constypeP1mCappuccino, constypeP1mRTD, r18br1, r18br2, r18br3, r18br4, r18br5, r18br6, r15_1, r15_2, r15_3, r15_4, r15_5, r15_6, r15_7, DURATIONWATCHINGTV, CurconABC, CurconABCSusu, CurconGoodDay, CurconIndocafe, CurconKapalApi, CurconKASpecial, CurconLuwak, CurconTorabika, CurconToraCappucino, CurconToraDuo, CurconToraMoka, CurconToraSusu, CurconTorabikaCreamyLatte, ConLMABC, ConLMABCSusu, ConLMGoodDay, ConLMIndocafe, ConLMKapalApi, ConLMKASpecial, ConLMLuwak, ConLMTorabika, ConLMToraCappucino, ConLMToraDuo, ConLMToraMoka, ConLMToraSusu, ConLMTorabikaCreamyLatte, Black_Coffee_Fav_Br, Black_Coffee_Bumo, Black_Coffee_PreBumo, Milk_Coffee_Fav_Br, Milk_Coffee_Bumo, Milk_Coffee_PreBumo, Mocca_Coffee_Fav_Br, Mocca_Coffee_Bumo, Mocca_Coffee_PreBumo, White_Coffee_Fav_Br, White_Coffee_Bumo, White_Coffee_PreBumo, Cappucinno_Fav_Br, Cappucinno_Bumo, Cappucinno_PreBumo, r27h1_34, r27h1_35, r27h1_23, r27h1_10, r27h1_25, r27h1_30, r27h3_34, r27h3_35, r27h3_23, r27h3_10, r27h3_25, r27h3_30, r27h4_34, r27h4_35, r27h4_23, r27h4_10, r27h4_25, r27h4_30, r27h7_34, r27h7_35, r27h7_23, r27h7_10, r27h7_25, r27h7_30, r27h9_34, r27h9_35, r27h9_23, r27h9_10, r27h9_25, r27h9_30, BrTomNet, BrSpontABC, BrSpontGoodDay, BrSpontKapalApi, BrSpontLuwak, BrSpontTorabika, BrAidABC, BrAidGoodDay, BrAidKapalApi, BrAidLuwak, BrAidTorabika, AdTomNet, AdSpontABC, AdSpontGoodDay, AdSpontKapalApi, AdSpontLuwak, AdSpontTorabika, AdAidABC, AdAidGoodDay, AdAidKapalApi, AdAidLuwak, AdAidTorabika, ConLMABC, ConLMGoodDay, ConLMKapalApi, ConLMLuwak, ConLMTorabika, CurconABC, CurconGoodDay, CurconKapalApi, CurconLuwak, CurconTorabika, Net_Fav, Net_Bumo, Net_PreBumo, r20_1, r20_2, r20_3, r20_4, r20_5, r20_6, BrSpontToraCafe_Volcano_Chocomelt, BrSpontTora_Cafe_Caramelove, BrSpontTora_Cafe_tidak_spesifik, BrAidToraCafe_Volcano_Chocomelt, BrAidTora_Cafe_Caramelove, BrAidTora_Cafe_tidak_spesifik, AdSpontToraCafe_Volcano_Chocomelt, AdSpontTora_Cafe_Caramelove, AdSpontTora_Cafe_tidak_spesifik, AdAidToraCafe_Volcano_Chocomelt, AdAidTora_Cafe_Caramelove, AdAidTora_Cafe_tidak_spesifik, CurconToraCafe_Volcano_Chocomelt, CurconTora_Cafe_Caramelove, CurconTora_Cafe_tidak_spesifik, ConL3MToraCafe_Volcano_Chocomelt, ConL3MTora_Cafe_Caramelove, ConL3MTora_Cafe_tidak_spesifik, BrSpontToraCafe, BrAidToraCafe, AdSpontToraCafe, AdAidToraCafe, ConsL3MToraCafe, CurconToraCafe, BrSpont_GoodDayCapucino, BrAid_GoodDayCapucino, AdSpont_GoodDayCapucino, AdAid_GoodDayCapucino, Curcon_GoodDayCapucino, ConsL3M_GoodDayCapucino, BrSpontTora_Kopi_Susu_Esp, BrAid_Tora_Kopi_Susu_Esp, AdSpont_Tora_Kopi_Susu_Esp, AdAid_Tora_Kopi_Susu_Esp, Curcon_Tora_Kopi_Susu_Esp, ConsL3M_Tora_Kopi_Susu_Esp, ConP1WABC, ConP1WABCSusu, ConP1WGoodDay, ConP1WIndocafe, ConP1WKapalApi, ConP1WKASpecial, ConP1WLuwak, ConP1WTorabika, ConP1WToraCappucino, ConP1WToraDuo, ConP1WToraMoka, ConP1WToraSusu, ConP1WTorabikaCreamyLatte, ConP1WToraCafe_Volcano_Chocomelt, ConP1WTora_Cafe_Caramelove, ConP1WTora_Cafe_tidak_spesifik, ConsP1W_Tora_Kopi_Susu_Esp, ConsP1W_GoodDayCapucino, ConP1DABC, ConP1DABCSusu, ConP1DGoodDay, ConP1DIndocafe, ConP1DKapalApi, ConP1DKASpecial, ConP1DLuwak, ConP1DTorabika, ConP1DToraCappucino, ConP1DToraDuo, ConP1DToraMoka, ConP1DToraSusu, ConP1DTorabikaCreamyLatte, ConP1DToraCafe_Volcano_Chocomelt, ConP1DTora_Cafe_Caramelove, ConP1DTora_Cafe_tidak_spesifik, ConsP1D_Tora_Kopi_Susu_Esp, ConsP1D_GoodDayCapucino, ConsP1DToraCafe, ConsP1WToraCafe, BrSpontTOP_Coffee_Cappuccino, BrAid_TOP_Coffee_Cappuccino, AdSpont_TOP_Coffee_Cappuccino, AdAid_TOP_Coffee_Cappuccino, Curcon_TOP_Coffee_Cappuccino, ConsL3M_TOP_Coffee_Cappuccino, ConsP1DTOP_Coffee_Cappuccino, ConsP1wTOP_Coffee_Cappuccino);
                    }
                }

            }
        }



        private static string getYear(string SurveyPeriod)
        {
            string[] date = SurveyPeriod.Split('-');
            return date[0];
        }

        private static string find_UserId(int SurveyID, string SurveyPeriod, string u_id)
        {
            string sum = "";
            string[] date = SurveyPeriod.Split('-');
            foreach (string d in date)
            {
                sum = sum + d;

            }
            return u_id + SurveyID + sum;
        }
    }

    class Insertion_m2coffee
    {
        SqlConnection connection = new SqlConnection("Data Source=52.74.59.117;Initial Catalog=ClueboxMobile;Persist Security Info=True;User ID=sa;Password=ClueBox123*;");
        internal void insert_Dashboard_variable_values(string VARIABLE_NAME, string VARIABLE_NAME_SUB_NAME, string VARIABLE_ID, string VARIABLE_VALUE, string VARIABLE_NAME_QUESTION, int SurveyID, string country, string BASE_VARIABLE_NAME, string SurveyPeriod)
        {
            connection.Open();
            SqlCommand cmd = new SqlCommand("INSERT INTO DashboardSPS_Variable_Values (VARIABLE_NAME,VARIABLE_NAME_SUB_NAME,VARIABLE_ID,VARIABLE_VALUE,VARIABLE_NAME_QUESTION,SURVEY_ID,SURVEY_COUNTRY,BASE_VARIABLE_NAME,SURVEY_PERIOD) " + "VALUES('" + VARIABLE_NAME + "','" + VARIABLE_NAME_SUB_NAME + "','" + VARIABLE_ID + "','" + VARIABLE_VALUE + "','" + VARIABLE_NAME_QUESTION + "','" + SurveyID + "','" + country + "','" + BASE_VARIABLE_NAME + "','" + SurveyPeriod + "')", connection);
            try
            {


                cmd.ExecuteNonQuery();
                Console.WriteLine("Dashboard variable values inserted successfully");

                connection.Close();



            }
            catch (Exception)
            {

                Console.WriteLine("Exception occured");
                Console.ReadLine();
            }

        }

        internal void insert_Dashboard_values(string userID, string Gender, string MaritalStatus, string country, int SurveyID, string SurveyPeriod, decimal Weight, string Location, string AgeGroup, string Ses, string Period, string year, string BrTom, string BrSpontABC, string BrSpontGoodDay, string BrSpontABCSusu, string BrSpontIndocafe, string BrSpontKapalApi, string BrSpontKASpecial, string BrSpontLuwak, string BrSpontTorabika, string BrSpontToraCappucino, string BrSpontToraDuo, string BrSpontToraMoka, string BrSpontToraSusu, string BrSpontTorabikaCreamyLatte, string BrAidABC, string BrAidABCSusu, string BrAidGoodDay, string BrAidIndocafe, string BrAidKapalApi, string BrAidKASpecial, string BrAidLuwak, string BrAidTorabika, string BrAidToraCappucino, string BrAidToraDuo, string BrAidToraMoka, string BrAidToraSusu, string BrAidTorabikaCreamyLatte, string AdTom, string AdSpontABC, string AdSpontABCSusu, string AdSpontGoodDay, string AdSpontIndocafe, string AdSpontKapalApi, string AdSpontKASpecial, string AdSpontLuwak, string AdSpontTorabika, string AdSpontToraCappucino, string AdSpontToraDuo, string AdSpontToraMoka, string AdSpontToraSusu, string AdSpontTorabikaCreamyLatte, string AdAidABC, string AdAidABCSusu, string AdAidGoodDay, string AdAidIndocafe, string AdAidKapalApi, string AdAidKASpecial, string AdAidLuwak, string AdAidTorabika, string AdAidToraCappucino, string AdAidToraDuo, string AdAidToraMoka, string AdAidToraSusu, string AdAidTorabikaCreamyLatte, string FavouriteBrand, string Bumo, string PrBumo, string RBCRTDTea, string RBCMineralwater, string RBCRTDJuice, string RBCCarbonatedRTD, string RBCRTDMilk, string RBCRTDCoffee, string RBCEnergydrink, string RBCVitamindrink, string RBCIsotonicdrink, string RBCYoghurtdrink, string RBCTeabag, string RBCInstantcoffee, string RBCInstantcoffeesachet, string RBCR_and_Gcoffeesachet, string RBCNone, string constypeP1mBlack, string constypeP1mmilk, string constypeP1mmocha, string constypeP1mWhite, string constypeP1mCappuccino, string constypeP1mRTD, string r18br1, string r18br2, string r18br3, string r18br4, string r18br5, string r18br6, string r15_1, string r15_2, string r15_3, string r15_4, string r15_5, string r15_6, string r15_7, string DURATIONWATCHINGTV, string CurconABC, string CurconABCSusu, string CurconGoodDay, string CurconIndocafe, string CurconKapalApi, string CurconKASpecial, string CurconLuwak, string CurconTorabika, string CurconToraCappucino, string CurconToraDuo, string CurconToraMoka, string CurconToraSusu, string CurconTorabikaCreamyLatte, string ConLMABC, string ConLMABCSusu, string ConLMGoodDay, string ConLMIndocafe, string ConLMKapalApi, string ConLMKASpecial, string ConLMLuwak, string ConLMTorabika, string ConLMToraCappucino, string ConLMToraDuo, string ConLMToraMoka, string ConLMToraSusu, string ConLMTorabikaCreamyLatte, string Black_Coffee_Fav_Br, string Black_Coffee_Bumo, string Black_Coffee_PreBumo, string Milk_Coffee_Fav_Br, string Milk_Coffee_Bumo, string Milk_Coffee_PreBumo, string Mocca_Coffee_Fav_Br, string Mocca_Coffee_Bumo, string Mocca_Coffee_PreBumo, string White_Coffee_Fav_Br, string White_Coffee_Bumo, string White_Coffee_PreBumo, string Cappucinno_Fav_Br, string Cappucinno_Bumo, string Cappucinno_PreBumo, string r27h1_34, string r27h1_35, string r27h1_23, string r27h1_10, string r27h1_25, string r27h1_30, string r27h3_34, string r27h3_35, string r27h3_23, string r27h3_10, string r27h3_25, string r27h3_30, string r27h4_34, string r27h4_35, string r27h4_23, string r27h4_10, string r27h4_25, string r27h4_30, string r27h7_34, string r27h7_35, string r27h7_23, string r27h7_10, string r27h7_25, string r27h7_30, string r27h9_34, string r27h9_35, string r27h9_23, string r27h9_10, string r27h9_25, string r27h9_30, string BrTomNet, string r21b_1, string r21b_55, string r21b_101, string r21b_133, string r21b_160, string r24a_1, string r24a_55, string r24a_101, string r24a_133, string r24a_160, string AdTomNet, string r21e_1, string r21e_55, string r21e_101, string r21e_133, string r21e_160, string r24b_1, string r24b_55, string r24b_101, string r24b_133, string r24b_160, string r25a_1, string r25a_55, string r25a_101, string r25a_133, string r25a_160, string r25b_1, string r25b_55, string r25b_101, string r25b_133, string r25b_160, string Net_Fav, string Net_Bumo, string Net_PreBumo, string r20_1, string r20_2, string r20_3, string r20_4, string r20_5, string r20_6, string BrSpontToraCafe_Volcano_Chocomelt, string BrSpontTora_Cafe_Caramelove, string BrSpontTora_Cafe_tidak_spesifik, string BrAidToraCafe_Volcano_Chocomelt, string BrAidTora_Cafe_Caramelove, string BrAidTora_Cafe_tidak_spesifik, string AdSpontToraCafe_Volcano_Chocomelt, string AdSpontTora_Cafe_Caramelove, string AdSpontTora_Cafe_tidak_spesifik, string AdAidToraCafe_Volcano_Chocomelt, string AdAidTora_Cafe_Caramelove, string AdAidTora_Cafe_tidak_spesifik, string CurconToraCafe_Volcano_Chocomelt, string CurconTora_Cafe_Caramelove, string CurconTora_Cafe_tidak_spesifik, string ConL3MToraCafe_Volcano_Chocomelt, string ConL3MTora_Cafe_Caramelove, string ConL3MTora_Cafe_tidak_spesifik, string BrSpontToraCafe, string BrAidToraCafe, string AdSpontToraCafe, string AdAidToraCafe, string ConsL3MToraCafe, string CurconToraCafe, string BrSpont_GoodDayCapucino, string BrAid_GoodDayCapucino, string AdSpont_GoodDayCapucino, string AdAid_GoodDayCapucino, string Curcon_GoodDayCapucino, string ConsL3M_GoodDayCapucino, string BrSpontTora_Kopi_Susu_Esp, string BrAid_Tora_Kopi_Susu_Esp, string AdSpont_Tora_Kopi_Susu_Esp, string AdAid_Tora_Kopi_Susu_Esp, string Curcon_Tora_Kopi_Susu_Esp, string ConsL3M_Tora_Kopi_Susu_Esp, string ConP1WABC, string ConP1WABCSusu, string ConP1WGoodDay, string ConP1WIndocafe, string ConP1WKapalApi, string ConP1WKASpecial, string ConP1WLuwak, string ConP1WTorabika, string ConP1WToraCappucino, string ConP1WToraDuo, string ConP1WToraMoka, string ConP1WToraSusu, string ConP1WTorabikaCreamyLatte, string ConP1WToraCafe_Volcano_Chocomelt, string ConP1WTora_Cafe_Caramelove, string ConP1WTora_Cafe_tidak_spesifik, string ConsP1W_Tora_Kopi_Susu_Esp, string ConsP1W_GoodDayCapucino, string ConP1DABC, string ConP1DABCSusu, string ConP1DGoodDay, string ConP1DIndocafe, string ConP1DKapalApi, string ConP1DKASpecial, string ConP1DLuwak, string ConP1DTorabika, string ConP1DToraCappucino, string ConP1DToraDuo, string ConP1DToraMoka, string ConP1DToraSusu, string ConP1DTorabikaCreamyLatte, string ConP1DToraCafe_Volcano_Chocomelt, string ConP1DTora_Cafe_Caramelove, string ConP1DTora_Cafe_tidak_spesifik, string ConsP1D_Tora_Kopi_Susu_Esp, string ConsP1D_GoodDayCapucino, string ConsP1DToraCafe, string ConsP1WToraCafe, string BrSpontTOP_Coffee_Cappuccino, string BrAid_TOP_Coffee_Cappuccino, string AdSpont_TOP_Coffee_Cappuccino, string AdAid_TOP_Coffee_Cappuccino, string Curcon_TOP_Coffee_Cappuccino, string ConsL3M_TOP_Coffee_Cappuccino, string ConsP1DTOP_Coffee_Cappuccino, string ConsP1wTOP_Coffee_Cappuccino)
        {
            int i;
            connection.Open();
            //SqlConnection connection = new SqlConnection("Data Source=52.74.59.117;Initial Catalog=ClueboxMobile;Persist Security Info=True;User ID=sa;Password=ClueBox123*;");
            SqlCommand cmd = new SqlCommand("INSERT INTO DashboardFlatTabJava_m2Coffee_temp (UserID,Gender,MaritalStatus,Country,SurveyID,AttendedOn,Weight,Location,AgeGroup,Ses,Period,Yearly,BrTom,BrSpontABC,BrSpontABCSusu,BrSpontGoodDay,BrSpontIndocafe,BrSpontKapalApi,BrSpontKASpecial,BrSpontLuwak,BrSpontTorabika,BrSpontToraCappucino,BrSpontToraDuo,BrSpontToraMoka,BrSpontToraSusu,BrSpontTorabikaCreamyLatte,BrAidABC,BrAidABCSusu,BrAidGoodDay,BrAidIndocafe,BrAidKapalApi,BrAidKASpecial,BrAidLuwak,BrAidTorabika,BrAidToraCappucino,BrAidToraDuo,BrAidToraMoka,BrAidToraSusu,BrAidTorabikaCreamyLatte,AdTom,AdSpontABC,AdSpontABCSusu,AdSpontGoodDay,AdSpontIndocafe,AdSpontKapalApi,AdSpontKASpecial,AdSpontLuwak,AdSpontTorabika,AdSpontToraCappucino,AdSpontToraDuo,AdSpontToraMoka,AdSpontToraSusu,AdSpontTorabikaCreamyLatte,AdAidABC,AdAidABCSusu,AdAidGoodDay,AdAidIndocafe,AdAidKapalApi,AdAidKASpecial,AdAidLuwak,AdAidTorabika,AdAidToraCappucino,AdAidToraDuo,AdAidToraMoka,AdAidToraSusu,AdAidTorabikaCreamyLatte,FavouriteBrand,Bumo,PrBumo,RBCRTDTea,RBCMineralwater,RBCRTDJuice,RBCCarbonatedRTD,RBCRTDMilk,RBCRTDCoffee,RBCEnergydrink,RBCVitamindrink,RBCIsotonicdrink,RBCYoghurtdrink,RBCTeabag,RBCInstantcoffee,RBCInstantcoffeesachet,RBCR_and_Gcoffeesachet,RBCNone,constypeP1mBlack,constypeP1mmilk,constypeP1mmocha,constypeP1mWhite,constypeP1mCappuccino,constypeP1mRTD,r18br1,r18br2,r18br3,r18br4,r18br5,r18br6,r15_1,r15_2,r15_3,r15_4,r15_5,r15_6,r15_7,DURATIONWATCHINGTV,CurconABC,CurconABCSusu,CurconGoodDay,CurconIndocafe,CurconKapalApi,CurconKASpecial,CurconLuwak,CurconTorabika,CurconToraCappucino,CurconToraDuo,CurconToraMoka,CurconToraSusu,CurconTorabikaCreamyLatte,ConL3MABC,ConL3MABCSusu,ConL3MGoodDay,ConL3MIndocafe,ConL3MKapalApi,ConL3MKASpecial,ConL3MLuwak,ConL3MTorabika,ConL3MToraCappucino,ConL3MToraDuo,ConL3MToraMoka,ConL3MToraSusu,ConL3MTorabikaCreamyLatte,Black_Coffee,Black_Coffee_Bumo,Black_Coffee_PreBumo,Milk_Coffee,Milk_Coffee_Bumo,Milk_Coffee_PreBumo,Mocca_Coffee,Mocca_Coffee_Bumo,Mocca_Coffee_PreBumo,White_Coffee,White_Coffee_Bumo,White_Coffee_PreBumo,Cappucinno,Cappucinno_Bumo,Cappucinno_PreBumo,r27h1_34,r27h1_35,r27h1_23,r27h1_10,r27h1_25,r27h1_30,r27h3_34,r27h3_35,r27h3_23,r27h3_10,r27h3_25,r27h3_30,r27h4_34,r27h4_35,r27h4_23,r27h4_10,r27h4_25,r27h4_30,r27h7_34,r27h7_35,r27h7_23,r27h7_10,r27h7_25,r27h7_30,r27h9_34,r27h9_35,r27h9_23,r27h9_10,r27h9_25,r27h9_30,BrTomNet,AdTomNet,Net_Fav,Net_Bumo,Net_PreBumo,BrTom_BK,AdTom_BK,BUMO_BK,Pre_BUMO_BK,FavouriteBrand_BK,BrTomNet_BK,AdTomNet_BK,Net_Fav_BK,Net_Bumo_BK,Net_PreBumo_BK,Black_Coffee_Pannel,Milk_Coffee_Pannel,Mocca_Coffee_Pannel,White_Coffee_Pannel,Cappucinno_Pannel,RTD_Coffee_Pannel,Black_Coffee_Fav_Br,Milk_Coffee_Fav_Br,Mocca_Coffee_Fav_Br,White_Coffee_Fav_Br,Cappucinno_Fav_Br,BrSpontToraCafe_Volcano_Chocomelt,BrSpontTora_Cafe_Caramelove,BrSpontTora_Cafe_tidak_spesifik,BrAidToraCafe_Volcano_Chocomelt,BrAidTora_Cafe_Caramelove,BrAidTora_Cafe_tidak_spesifik,AdSpontToraCafe_Volcano_Chocomelt,AdSpontTora_Cafe_Caramelove,AdSpontTora_Cafe_tidak_spesifik,AdAidToraCafe_Volcano_Chocomelt,AdAidTora_Cafe_Caramelove,AdAidTora_Cafe_tidak_spesifik,CurconToraCafe_Volcano_Chocomelt,CurconTora_Cafe_Caramelove,CurconTora_Cafe_tidak_spesifik,ConL3MToraCafe_Volcano_Chocomelt,ConL3MTora_Cafe_Caramelove,ConL3MTora_Cafe_tidak_spesifik,BrSpontToraCafe,BrAidToraCafe,AdSpontToraCafe,AdAidToraCafe,ConsL3MToraCafe,CurconToraCafe,BrSpont_GoodDayCapucino,BrAid_GoodDayCapucino,AdSpont_GoodDayCapucino,AdAid_GoodDayCapucino,Curcon_GoodDayCapucino,ConsL3M_GoodDayCapucino,BrSpontTora_Kopi_Susu_Esp,BrAid_Tora_Kopi_Susu_Esp ,AdSpont_Tora_Kopi_Susu_Esp,AdAid_Tora_Kopi_Susu_Esp,Curcon_Tora_Kopi_Susu_Esp,ConsL3M_Tora_Kopi_Susu_Esp,ConP1WABC ,ConP1WABCSusu ,ConP1WGoodDay ,ConP1WIndocafe ,ConP1WKapalApi ,ConP1WKASpecial ,ConP1WLuwak ,ConP1WTorabika ,ConP1WToraCappucino ,ConP1WToraDuo ,ConP1WToraMoka ,ConP1WToraSusu ,ConP1WTorabikaCreamyLatte ,ConP1WToraCafe_Volcano_Chocomelt ,ConP1WTora_Cafe_Caramelove ,ConP1WTora_Cafe_tidak_spesifik ,ConsP1W_Tora_Kopi_Susu_Esp ,ConsP1W_GoodDayCapucino ,ConP1DABC ,ConP1DABCSusu ,ConP1DGoodDay ,ConP1DIndocafe ,ConP1DKapalApi ,ConP1DKASpecial ,ConP1DLuwak ,ConP1DTorabika ,ConP1DToraCappucino ,ConP1DToraDuo ,ConP1DToraMoka ,ConP1DToraSusu ,ConP1DTorabikaCreamyLatte ,ConP1DToraCafe_Volcano_Chocomelt ,ConP1DTora_Cafe_Caramelove ,ConP1DTora_Cafe_tidak_spesifik ,ConsP1D_Tora_Kopi_Susu_Esp ,ConsP1D_GoodDayCapucino ,ConsP1DToraCafe ,ConsP1WToraCafe,BrSpontTOP_Coffee_Cappuccino,BrAid_TOP_Coffee_Cappuccino,AdSpont_TOP_Coffee_Cappuccino,AdAid_TOP_Coffee_Cappuccino,Curcon_TOP_Coffee_Cappuccino,ConsL3M_TOP_Coffee_Cappuccino,ConsP1DTOP_Coffee_Cappuccino,ConsP1wTOP_Coffee_Cappuccino ) " + "VALUES('" + userID + "','" + Gender + "','" + MaritalStatus + "','" + country + "','" + SurveyID + "','" + SurveyPeriod + "','" + Weight + "','" + Location + "','" + AgeGroup + "','" + Ses + "','" + Period + "','" + year + "','" + BrTom + "','" + BrSpontABC + "','" + BrSpontGoodDay + "','" + BrSpontABCSusu + "','" + BrSpontIndocafe + "','" + BrSpontKapalApi + "','" + BrSpontKASpecial + "','" + BrSpontLuwak + "','" + BrSpontTorabika + "','" + BrSpontToraCappucino + "','" + BrSpontToraDuo + "','" + BrSpontToraMoka + "','" + BrSpontToraSusu + "','" + BrSpontTorabikaCreamyLatte + "','" + BrAidABC + "','" + BrAidABCSusu + "','" + BrAidGoodDay + "','" + BrAidIndocafe + "','" + BrAidKapalApi + "','" + BrAidKASpecial + "','" + BrAidLuwak + "','" + BrAidTorabika + "','" + BrAidToraCappucino + "','" + BrAidToraDuo + "','" + BrAidToraMoka + "','" + BrAidToraSusu + "','" + BrAidTorabikaCreamyLatte + "','" + AdTom + "','" + AdSpontABC + "','" + AdSpontABCSusu + "','" + AdSpontGoodDay + "','" + AdSpontIndocafe + "','" + AdSpontKapalApi + "','" + AdSpontKASpecial + "','" + AdSpontLuwak + "','" + AdSpontTorabika + "','" + AdSpontToraCappucino + "','" + AdSpontToraDuo + "','" + AdSpontToraMoka + "','" + AdSpontToraSusu + "','" + AdSpontTorabikaCreamyLatte + "','" + AdAidABC + "','" + AdAidABCSusu + "','" + AdAidGoodDay + "','" + AdAidIndocafe + "','" + AdAidKapalApi + "','" + AdAidKASpecial + "','" + AdAidLuwak + "','" + AdAidTorabika + "','" + AdAidToraCappucino + "','" + AdAidToraDuo + "','" + AdAidToraMoka + "','" + AdAidToraSusu + "','" + AdAidTorabikaCreamyLatte + "','" + FavouriteBrand + "','" + Bumo + "','" + PrBumo + "','" + RBCRTDTea + "','" + RBCMineralwater + "','" + RBCRTDJuice + "','" + RBCCarbonatedRTD + "','" + RBCRTDMilk + "','" + RBCRTDCoffee + "','" + RBCEnergydrink + "','" + RBCVitamindrink + "','" + RBCIsotonicdrink + "','" + RBCYoghurtdrink + "','" + RBCTeabag + "','" + RBCInstantcoffee + "','" + RBCInstantcoffeesachet + "','" + RBCR_and_Gcoffeesachet + "','" + RBCNone + "','" + constypeP1mBlack + "','" + constypeP1mmilk + "','" + constypeP1mmocha + "','" + constypeP1mWhite + "','" + constypeP1mCappuccino + "','" + constypeP1mRTD + "','" + r18br1 + "','" + r18br2 + "','" + r18br3 + "','" + r18br4 + "','" + r18br5 + "','" + r18br6 + "','" + r15_1 + "','" + r15_2 + "','" + r15_3 + "','" + r15_4 + "','" + r15_5 + "','" + r15_6 + "','" + r15_7 + "','" + DURATIONWATCHINGTV + "','" + CurconABC + "','" + CurconABCSusu + "','" + CurconGoodDay + "','" + CurconIndocafe + "','" + CurconKapalApi + "','" + CurconKASpecial + "','" + CurconLuwak + "','" + CurconTorabika + "','" + CurconToraCappucino + "','" + CurconToraDuo + "','" + CurconToraMoka + "','" + CurconToraSusu + "','" + CurconTorabikaCreamyLatte + "','" + ConLMABC + "','" + ConLMABCSusu + "','" + ConLMGoodDay + "','" + ConLMIndocafe + "','" + ConLMKapalApi + "','" + ConLMKASpecial + "','" + ConLMLuwak + "','" + ConLMTorabika + "','" + ConLMToraCappucino + "','" + ConLMToraDuo + "','" + ConLMToraMoka + "','" + ConLMToraSusu + "','" + ConLMTorabikaCreamyLatte + "','" + Black_Coffee_Fav_Br + "','" + Black_Coffee_Bumo + "','" + Black_Coffee_PreBumo + "','" + Milk_Coffee_Fav_Br + "','" + Milk_Coffee_Bumo + "','" + Milk_Coffee_PreBumo + "','" + Mocca_Coffee_Fav_Br + "','" + Mocca_Coffee_Bumo + "','" + Mocca_Coffee_PreBumo + "','" + White_Coffee_Fav_Br + "','" + White_Coffee_Bumo + "','" + White_Coffee_PreBumo + "','" + Cappucinno_Fav_Br + "','" + Cappucinno_Bumo + "','" + Cappucinno_PreBumo + "','" + r27h1_34 + "','" + r27h1_35 + "','" + r27h1_23 + "','" + r27h1_10 + "','" + r27h1_25 + "','" + r27h1_30 + "','" + r27h3_34 + "','" + r27h3_35 + "','" + r27h3_23 + "','" + r27h3_10 + "','" + r27h3_25 + "','" + r27h3_30 + "','" + r27h4_34 + "','" + r27h4_35 + "','" + r27h4_23 + "','" + r27h4_10 + "','" + r27h4_25 + "','" + r27h4_30 + "','" + r27h7_34 + "','" + r27h7_35 + "','" + r27h7_23 + "','" + r27h7_10 + "','" + r27h7_25 + "','" + r27h7_30 + "','" + r27h9_34 + "','" + r27h9_35 + "','" + r27h9_23 + "','" + r27h9_10 + "','" + r27h9_25 + "','" + r27h9_30 + "','" + BrTomNet + "','" + AdTomNet + "','" + Net_Fav + "','" + Net_Bumo + "','" + Net_PreBumo + "','" + BrTom + "','" + AdTom + "','" + Bumo + "','" + PrBumo + "','" + FavouriteBrand + "','" + BrTomNet + "','" + AdTomNet + "','" + Net_Fav + "','" + Net_Bumo + "','" + Net_PreBumo + "','" + r20_1 + "','" + r20_2 + "','" + r20_3 + "','" + r20_4 + "','" + r20_5 + "','" + r20_6 + "','" + Black_Coffee_Fav_Br + "','" + Milk_Coffee_Fav_Br + "','" + Mocca_Coffee_Fav_Br + "','" + White_Coffee_Fav_Br + "','" + Cappucinno_Fav_Br + "','" + BrSpontToraCafe_Volcano_Chocomelt + "','" + BrSpontTora_Cafe_Caramelove + "','" + BrSpontTora_Cafe_tidak_spesifik + "','" + BrAidToraCafe_Volcano_Chocomelt + "','" + BrAidTora_Cafe_Caramelove + "','" + BrAidTora_Cafe_tidak_spesifik + "','" + AdSpontToraCafe_Volcano_Chocomelt + "','" + AdSpontTora_Cafe_Caramelove + "','" + AdSpontTora_Cafe_tidak_spesifik + "','" + AdAidToraCafe_Volcano_Chocomelt + "','" + AdAidTora_Cafe_Caramelove + "','" + AdAidTora_Cafe_tidak_spesifik + "','" + CurconToraCafe_Volcano_Chocomelt + "','" + CurconTora_Cafe_Caramelove + "','" + CurconTora_Cafe_tidak_spesifik + "','" + ConL3MToraCafe_Volcano_Chocomelt + "','" + ConL3MTora_Cafe_Caramelove + "','" + ConL3MTora_Cafe_tidak_spesifik + "','" + BrSpontToraCafe + "','" + BrAidToraCafe + "','" + AdSpontToraCafe + "','" + AdAidToraCafe + "','" + ConsL3MToraCafe + "','" + CurconToraCafe + "','" + BrSpont_GoodDayCapucino + "','" + BrAid_GoodDayCapucino + "','" + AdSpont_GoodDayCapucino + "','" + AdAid_GoodDayCapucino + "','" + Curcon_GoodDayCapucino + "','" + ConsL3M_GoodDayCapucino + "','" + BrSpontTora_Kopi_Susu_Esp + "','" + BrAid_Tora_Kopi_Susu_Esp + "','" + AdSpont_Tora_Kopi_Susu_Esp + "','" + AdAid_Tora_Kopi_Susu_Esp + "','" + Curcon_Tora_Kopi_Susu_Esp + "','" + ConsL3M_Tora_Kopi_Susu_Esp + "' ,'" + ConP1WABC + "','" + ConP1WABCSusu + "','" + ConP1WGoodDay + "','" + ConP1WIndocafe + "','" + ConP1WKapalApi + "','" + ConP1WKASpecial + "','" + ConP1WLuwak + "','" + ConP1WTorabika + "','" + ConP1WToraCappucino + "','" + ConP1WToraDuo + "','" + ConP1WToraMoka + "','" + ConP1WToraSusu + "','" + ConP1WTorabikaCreamyLatte + "','" + ConP1WToraCafe_Volcano_Chocomelt + "','" + ConP1WTora_Cafe_Caramelove + "','" + ConP1WTora_Cafe_tidak_spesifik + "','" + ConsP1W_Tora_Kopi_Susu_Esp + "','" + ConsP1W_GoodDayCapucino + "','" + ConP1DABC + "','" + ConP1DABCSusu + "','" + ConP1DGoodDay + "','" + ConP1DIndocafe + "','" + ConP1DKapalApi + "','" + ConP1DKASpecial + "','" + ConP1DLuwak + "','" + ConP1DTorabika + "','" + ConP1DToraCappucino + "','" + ConP1DToraDuo + "','" + ConP1DToraMoka + "','" + ConP1DToraSusu + "','" + ConP1DTorabikaCreamyLatte + "','" + ConP1DToraCafe_Volcano_Chocomelt + "','" + ConP1DTora_Cafe_Caramelove + "','" + ConP1DTora_Cafe_tidak_spesifik + "','" + ConsP1D_Tora_Kopi_Susu_Esp + "' ,'" + ConsP1D_GoodDayCapucino + "' ,'" + ConsP1DToraCafe + "' ,'" + ConsP1WToraCafe + "' ,'" + BrSpontTOP_Coffee_Cappuccino + "' ,'" + BrAid_TOP_Coffee_Cappuccino + "' ,'" + AdSpont_TOP_Coffee_Cappuccino + "' ,'" + AdAid_TOP_Coffee_Cappuccino + "' ,'" + Curcon_TOP_Coffee_Cappuccino + "' ,'" + ConsL3M_TOP_Coffee_Cappuccino + "' ,'" + ConsP1DTOP_Coffee_Cappuccino + "' ,'" + ConsP1wTOP_Coffee_Cappuccino + "' )", connection);
            try
            {
                cmd.ExecuteNonQuery();
                Console.WriteLine("Data inserted successfully");
                i = 1;
            }
            catch (Exception ex)
            {

                connection.Close();
                i = 0;
                Console.WriteLine("Exception occured" + ex.ToString());
                Console.WriteLine("Exception occured");

                Console.ReadLine();
            }
            connection.Close();
        }
    }
}

