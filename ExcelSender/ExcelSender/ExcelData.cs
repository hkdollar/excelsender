using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSender
{
    class ExcelData
    {
        //string Level0;
        //string Level1;
        //string Level2;
        //int Year;
        //int NumberOfDeathMen;
        //int NumberOfDeathWomen;
        //float RateOfDeathMen;
        //float RateOfDeathWomen;
        //float RateOfDeath;
        //int TotalPopulation;
        //int OldPopulation;
        //int NumberofDivorce;
        //float RateOfStress;

        public object Level0;
        public object Level1;
        public object Level2;
        public object Year;
        public object NumberOfDeathMen;
        public object NumberOfDeathWomen;
        public object RateOfDeathMen;
        public object RateOfDeathWomen;
        public object RateOfDeath;
        public object TotalPopulation;
        public object OldPopulation;
        public object NumberofDivorce;
        public object RateOfStress;


        //public ExcelData(string lv0, string lv1, string lv2, int yr, int nodeathmen, int nodeathwomen,  float ratedeathmen, float ratedeathwomen, float ratedeath,
        //            int totalpopulation, int oldpopulation, int nodivorce, float ratestress)
        //{
        //    Level0 = lv0;
        //    Level1 = lv1;
        //    Level2 = lv2;
        //    Year = yr;
        //    NumberOfDeathMen = nodeathmen;
        //    NumberOfDeathWomen = nodeathwomen;
        //    RateOfDeathMen = ratedeathmen;
        //    RateOfDeathWomen = ratedeathwomen;
        //    RateOfDeath = ratedeath;
        //    TotalPopulation = totalpopulation;
        //    OldPopulation = oldpopulation;
        //    NumberofDivorce = nodivorce;
        //    RateOfStress = ratestress;
        //}

        public ExcelData(object lv0, object lv1, object lv2, object yr, object nodeathmen, object nodeathwomen, object ratedeathmen, object ratedeathwomen, object ratedeath,
            object totalpopulation, object oldpopulation, object nodivorce, object ratestress)
        {
            Level0 = lv0;
            Level1 = lv1;
            Level2 = lv2;
            Year = yr;
            NumberOfDeathMen = nodeathmen;
            NumberOfDeathWomen = nodeathwomen;
            RateOfDeathMen = ratedeathmen;
            RateOfDeathWomen = ratedeathwomen;
            RateOfDeath = ratedeath;
            TotalPopulation = totalpopulation;
            OldPopulation = oldpopulation;
            NumberofDivorce = nodivorce;
            RateOfStress = ratestress;
        }

    }
}
