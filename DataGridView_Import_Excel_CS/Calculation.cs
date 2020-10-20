using System;
using System.Data;

namespace DataGridView_Import_Excel
{
    public class Calculation
    {
        //calculation of Q3
        public static void Q3calc(DataRow row, ref Saloon Q3_326)
        {
            if (row[6].ToString().Contains("FC"))
            {
                Q3_326.FCcount += 1;
                Q3_326.FCtime += Convert.ToInt16(row[7].ToString());
            }
            else if (row[6].ToString().Contains("FB"))
            {
                Q3_326.FBcount += 1;
                Q3_326.FBtime += Convert.ToInt16(row[7].ToString());
            }

            else if (row[6].ToString().Contains("RB"))
            {
                if (row[6].ToString().Contains("RB60"))
                {
                    Q3_326.RB60count += 1;
                    Q3_326.RB60time += Convert.ToInt16(row[7].ToString());
                }
                else if (row[6].ToString().Contains("RB40") || row[6].ToString().Contains("RB20"))
                {
                    if (row[6].ToString().Contains("RB40"))
                    {
                        Q3_326.RB40count += 1;
                        Q3_326.RB40time += Convert.ToInt16(row[7].ToString());
                    }
                    else if (row[6].ToString().Contains("RB20"))
                    {
                        Q3_326.RB20count += 1;
                        Q3_326.RB20time += Convert.ToInt16(row[7].ToString());
                    }
                }
            }
            else if (row[6].ToString().Contains("RC"))
            {
                if (row[6].ToString().Contains("RC40") || row[6].ToString().Contains("RC20"))
                {

                    Q3_326.RC40count += 1;
                    Q3_326.RC40time += Convert.ToInt16(row[7].ToString());

                }
                else if (row[6].ToString().Contains("RC60"))
                {
                    Q3_326.RC60count += 1;
                    Q3_326.RC60time += Convert.ToInt16(row[7].ToString());
                }
                else
                {
                    Q3_326.RCcount += 1;
                    Q3_326.RCtime += Convert.ToInt16(row[7].ToString());
                }

            }
        }

        //calculation SK38
        public static void SK38calc(DataRow row , ref Saloon Skoda)
        {
            if (row[6].ToString().ToUpper().Contains("FC"))
            {
                Skoda.FCcount += 1;
                Skoda.FCtime += Convert.ToInt16(row[7].ToString());
            }

            else if (row[6].ToString().ToUpper().Contains("FB"))
            {
                Skoda.FBcount += 1;
                Skoda.FBtime += Convert.ToInt16(row[7].ToString());
            }

            else if (row[6].ToString().ToUpper().Contains("RB"))
            {
                if (row[6].ToString().ToUpper().Contains("RB40"))
                {
                    Skoda.RB40count += 1;
                    Skoda.RB40time += Convert.ToInt16(row[7].ToString());
                }
                if (row[6].ToString().ToUpper().Contains("RB60"))
                {
                    Skoda.RB60count += 1;
                    Skoda.RB60time += Convert.ToInt16(row[7].ToString());
                }
            }

            else if (row[6].ToString().ToUpper().Contains("RC"))
            {
                Skoda.RCcount += 1;
                Skoda.RCtime += Convert.ToInt16(row[7].ToString());
            }
        }

        //calculation of DAIMLER BR223
        public static void BR223calc(DataRow row, ref Saloon BR223)
        {
            //Console.WriteLine(row[6].ToString());
            if (row[6].ToString().ToUpper().Contains("FAKI"))
            {
                BR223.FCcount += 1;
                BR223.FCtime += Convert.ToInt16(row[7].ToString());
            }

            else if (row[6].ToString().ToUpper().Contains("FALE"))
            {
                BR223.FBcount += 1;
                BR223.FBtime += Convert.ToInt16(row[7].ToString());
            }

            else if (row[6].ToString().ToUpper().Contains("FOLE"))
            {
                BR223.RBcount += 1;
                BR223.RBtime += Convert.ToInt16(row[7].ToString());
            }

            else if (row[6].ToString().ToUpper().Contains("FOKI"))
            {
                BR223.RCcount += 1;
                BR223.RCtime += Convert.ToInt16(row[7].ToString());
            }
            else
            {
                BR223.VSTcount += 1;
                BR223.VSTtime += Convert.ToInt16(row[7].ToString());
            }
        }

        //calculation G11
        public static void G11calc (DataRow row,ref Saloon G11)
        {
            if (row[6].ToString().ToUpper().Contains("FC"))
            {
                G11.FCcount += 1;
                G11.FCtime += Convert.ToInt16(row[7].ToString());
            }
            else if (row[6].ToString().ToUpper().Contains("FB"))
            {
                G11.FBcount += 1;
                G11.FBtime += Convert.ToInt16(row[7].ToString());
            }
            else if (row[6].ToString().ToUpper().Contains("RC"))
            {
                if (row[6].ToString().ToUpper().Contains("RC100"))
                {
                    G11.RC100count += 1;
                    G11.RC100time += Convert.ToInt16(row[7].ToString());
                }
                else
                {
                    G11.RC40count += 1;
                    G11.RC40time += Convert.ToInt16(row[7].ToString());
                }


            }
            else if (row[6].ToString().ToUpper().Contains("RB"))
            {
                G11.RBcount += 1;
                G11.RBtime += Convert.ToInt16(row[7].ToString());
            }
        }

        //calculation G3
        public static void G3calc (DataRow row, ref Saloon G3)
        {
            if (row[6].ToString().ToUpper().Contains("FC"))
            {
                G3.FCcount += 1;
                G3.FCtime += Convert.ToInt16(row[7].ToString());
            }
            else if (row[6].ToString().ToUpper().Contains("FB"))
            {
                G3.FBcount += 1;
                G3.FBtime += Convert.ToInt16(row[7].ToString());
            }
            else if (row[6].ToString().ToUpper().Contains("RC"))
            {
                if (row[6].ToString().ToUpper().Contains("RC100"))
                {
                    G3.RC100count += 1;
                    G3.RC100time += Convert.ToInt16(row[7].ToString());
                }
                else
                {
                    G3.RC40count += 1;
                    G3.RC40time += Convert.ToInt16(row[7].ToString());
                }


            }
            else if (row[6].ToString().ToUpper().Contains("RB"))
            {
                G3.RBcount += 1;
                G3.RBtime += Convert.ToInt16(row[7].ToString());
            }
        }


    }
}
