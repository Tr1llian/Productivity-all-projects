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
                if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                {
                    Q3_326.FCcount += 1;
                }
                Q3_326.FCtime += Convert.ToInt16(row[7].ToString());
            }
            else if (row[6].ToString().Contains("FB"))
            {
                if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                {
                    Q3_326.FBcount += 1;
                }
                Q3_326.FBtime += Convert.ToInt16(row[7].ToString());
            }

            else if (row[6].ToString().Contains("RB"))
            {
                if (row[6].ToString().Contains("RB60"))
                {
                    if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                    {
                        Q3_326.RB60count += 1;
                    }
                    Q3_326.RB60time += Convert.ToInt16(row[7].ToString());
                }
                else if (row[6].ToString().Contains("RB40") || row[6].ToString().Contains("RB20"))
                {
                    if (row[6].ToString().Contains("RB40"))
                    {
                        if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                        {
                            Q3_326.RB40count += 1;
                        }
                        Q3_326.RB40time += Convert.ToInt16(row[7].ToString());
                    }
                    else if (row[6].ToString().Contains("RB20"))
                    {
                        if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                        {
                            Q3_326.RB20count += 1;
                        }
                        Q3_326.RB20time += Convert.ToInt16(row[7].ToString());
                    }
                }
            }
            else if (row[6].ToString().Contains("RC"))
            {
                if (row[6].ToString().Contains("RC40") || row[6].ToString().Contains("RC20"))
                {
                    if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                    {
                        Q3_326.RC40count += 1;
                    }
                    Q3_326.RC40time += Convert.ToInt16(row[7].ToString());

                }
                else if (row[6].ToString().Contains("RC60"))
                {
                    if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                    {
                        Q3_326.RC60count += 1;
                    }
                    Q3_326.RC60time += Convert.ToInt16(row[7].ToString());
                }
                else
                {
                    if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                    {
                        Q3_326.RCcount += 1;
                    }
                    Q3_326.RCtime += Convert.ToInt16(row[7].ToString());
                }

            }
        }

        //calculation SK38
        public static void SK38calc(DataRow row , ref Saloon Skoda)
        {
            if (row[6].ToString().ToUpper().Contains("FC"))
            {
                if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                {
                    Skoda.FCcount += 1;
                }
                Skoda.FCtime += Convert.ToInt16(row[7].ToString());
            }

            else if (row[6].ToString().ToUpper().Contains("FB"))
            {
                if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                {
                    Skoda.FBcount += 1;
                }
                Skoda.FBtime += Convert.ToInt16(row[7].ToString());
            }

            else if (row[6].ToString().ToUpper().Contains("RB"))
            {
                if (row[6].ToString().ToUpper().Contains("RB40"))
                {
                    if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                    {
                        Skoda.RB40count += 1;
                    }
                    Skoda.RB40time += Convert.ToInt16(row[7].ToString());
                }
                if (row[6].ToString().ToUpper().Contains("RB60"))
                {
                    if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                    {
                        Skoda.RB60count += 1;
                    }
                    Skoda.RB60time += Convert.ToInt16(row[7].ToString());
                }
            }

            else if (row[6].ToString().ToUpper().Contains("RC"))
            {
                if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                {
                    Skoda.RCcount += 1;
                }
                Skoda.RCtime += Convert.ToInt16(row[7].ToString());
            }
        }

        //calculation of DAIMLER BR223
        public static void BR223calc(DataRow row, ref Saloon BR223)
        {
            if (row[6].ToString().ToUpper().Contains("VST")|| row[6].ToString().ToUpper().Contains("MITTE") || row[6].ToString().ToUpper().Contains("M-TEIL")|| row[6].ToString().Contains("Motorschutzabdeckung")|| row[6].ToString().Contains("Sichtschutz"))
            {
                BR223.VSTtime+= Convert.ToInt16(row[7].ToString());
                BR223.VSTcount++;
            }
                //Console.WriteLine(row[6].ToString());
                if (row[6].ToString().ToUpper().Contains("FAKI"))
            {
                if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                {
                    BR223.FCcount += 1;
                }
                BR223.FCtime += Convert.ToInt16(row[7].ToString());
            }

            else if (row[6].ToString().ToUpper().Contains("FALE"))
            {
                if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                {
                    BR223.FBcount += 1;
                }
                BR223.FBtime += Convert.ToInt16(row[7].ToString());
            }

            else if (row[6].ToString().ToUpper().Contains("FOLE"))
            {
                if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                {
                    if ((row[6].ToString().ToUpper().Contains("FSS")))
                    {
                        BR223.RB100count += 1;
                    }
                    else if ((row[6].ToString().ToUpper().Contains("FES")))
                    {
                        BR223.RB40count += 1;
                    }
                }
                if (row[6].ToString().ToUpper().Contains("FES") || (row[6].ToString().ToUpper().Contains("FSS")))
                {
                    BR223.RBtime += Convert.ToInt16(row[7].ToString());
                }
            }

            else if (row[6].ToString().ToUpper().Contains("FOKI"))
            {
                if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                {
                    if (row[6].ToString().ToUpper().Contains("FOKI") && (row[6].ToString().ToUpper().Contains("FSS")))
                    {
                        BR223.RC100count += 1;
                    }
                    if(row[6].ToString().ToUpper().Contains("FOKI") && (row[6].ToString().ToUpper().Contains("FES")))
                    {
                        BR223.RC40count += 1;
                    }

                }
                if (row[6].ToString().ToUpper().Contains("FES") || (row[6].ToString().ToUpper().Contains("FSS")))
                {
                    BR223.RCtime += Convert.ToInt16(row[7].ToString());
                }
            }
            
        }

        //calculation G11
        public static void G11calc (DataRow row,ref Saloon G11)
        {
            if (row[6].ToString().ToUpper().Contains("FC"))
            {
                if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                {
                    G11.FCcount += 1;
                }
                G11.FCtime += Convert.ToInt16(row[7].ToString());
            }
            else if (row[6].ToString().ToUpper().Contains("FB"))
            {
                if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                {
                    G11.FBcount += 1;
                }
                G11.FBtime += Convert.ToInt16(row[7].ToString());
            }
            else if (row[6].ToString().ToUpper().Contains("RC"))
            {
                if (row[6].ToString().ToUpper().Contains("RC100"))
                {
                    if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                    {
                        G11.RC100count += 1;
                    }
                    G11.RC100time += Convert.ToInt16(row[7].ToString());
                }
                else
                {
                    if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                    {
                        G11.RC40count += 1;
                    }
                    G11.RC40time += Convert.ToInt16(row[7].ToString());
                }


            }
            else if (row[6].ToString().ToUpper().Contains("RB"))
            {
                if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                {
                    G11.RBcount += 1;
                }
                G11.RBtime += Convert.ToInt16(row[7].ToString());
            }
        }

        //calculation G3
        public static void G3calc (DataRow row, ref Saloon G3)
        {
            if (row[6].ToString().ToUpper().Contains("FC"))
            {
                if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                {
                    G3.FCcount += 1;
                }
                G3.FCtime += Convert.ToInt16(row[7].ToString());
            }
            else if (row[6].ToString().ToUpper().Contains("FB"))
            {
                if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                {
                    G3.FBcount += 1;
                }
                G3.FBtime += Convert.ToInt16(row[7].ToString());
            }
            else if (row[6].ToString().ToUpper().Contains("RC"))
            {
                if (row[6].ToString().ToUpper().Contains("RC100"))
                {
                    if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                    {
                        G3.RC100count += 1;
                    }
                    G3.RC100time += Convert.ToInt16(row[7].ToString());
                }
                else
                {
                    if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                    {
                        G3.RC40count += 1;
                    }
                    G3.RC40time += Convert.ToInt16(row[7].ToString());
                }


            }
            else if (row[6].ToString().ToUpper().Contains("RB"))
            {
                if (Convert.ToInt32(row[3].ToString()) >= 5000000)
                {
                    G3.RBcount += 1;
                }
                G3.RBtime += Convert.ToInt16(row[7].ToString());
            }
        }

        //calculation BMWvoga
        public static void BMWvogacalc(Saloon G11, Saloon G3, ref Saloon BMWvoga)
        {
            BMWvoga.FBcount = G11.FBcount + G3.FBcount;
            BMWvoga.FCcount = G11.FCcount + G3.FCcount;
            BMWvoga.FBtime = G11.FBtime + G3.FBtime;
            BMWvoga.FCtime = G11.FCtime + G11.FCtime;
        }

        //calculation BMWhiga
        public static void BMWhigacalc(Saloon G11,ref Saloon BMWhiga)
        {
            BMWhiga.RBcount = G11.RBcount;
            BMWhiga.RBtime = G11.RBtime;
            BMWhiga.RCtime = G11.RC40time + G11.RC100time;
            BMWhiga.RCcount = G11.RC100count + G11.RC40count;
        }

    }
}
