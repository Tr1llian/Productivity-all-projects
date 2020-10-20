using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DataGridView_Import_Excel
{
    class Saloon
    {
        public string ProjectName;
        public int FCtime;
        public int FCcount;




        public int FBtime;
        public int FBcount;
       

        public int RCtime;
        public int RCcount;
       

        public int RC40time;
        public int RC40count;
        

        public int RC60time;
        public int RC60count;
       

        public int RB40time;
        public int RB40count;
       

        public int RB60time;
        public int RB60count;

        public int RB20time;
        public int RB20count;

        public int RBtime;
        public int RBcount;

        public int Coef = 1;

        public  Saloon(string c)
        {
            ProjectName = c;
            FCcount = FCtime = FBcount = FBtime = RCcount = RCtime = RBtime = RBcount=RC40count=RC40time=RC60count=RC60time=0;
        }

        public double AvgTime() 
        {
            double AllPcs= FCcount + FBcount + RCcount + RBcount + RB20count;
            double AllTime = FCtime + FBtime + RB20time + RBtime + RCtime;
            if (AllPcs == 0)
            {
                return 0;
            }
            else
            {
                return AllTime / AllPcs;
            }
        }

        public double GeneralCount()
        {
            //RBtime += RB40time + RB60time;
            //RBcount += RB60count + RB40count;
            return FCcount + FBcount + RCcount + RBcount + RB20count;
        }

        public double CompleteSaloons()
        {
            RBtime += RB40time + RB60time;
            RBcount += RB60count + RB40count;
            int[] arr = new int[4]; 
            arr[0] = Convert.ToInt32(FBcount/2);
            arr[1] = Convert.ToInt32(FCcount / 2);
            arr[2] = Convert.ToInt32(RBcount / 2);
            arr[3] = Convert.ToInt32(RCcount / 2);
            //arr[0] =Convert.ToInt32( PartTime(RB20time, RB20count) );

            return arr.Min();
        }

        public double PartTime(double  a, double b)
        {
            if (a != 0 && b != 0)
            {
                return a / b;
            }
            else return 0;
        }
        public double TimeSaloon()
        {
            RBtime += RB40time + RB60time;
            RBcount += RB60count + RB40count;

            if (RB20count != 0 && RB20time != 0)
            {

                if (FCtime != 0 && FBtime != 0 && RBtime != 0 && RCtime != 0)
                {
                    return 2 * (PartTime(FCtime,FCcount) + PartTime(FBtime,FBcount) + PartTime(RBtime,RBcount) + PartTime(RCtime,RCcount))+PartTime(RB20time,RB20count);
                }
                if (FCtime != 0 && FBtime != 0 && RBtime == 0 && RCtime == 0)
                {
                    return (2 * PartTime(FCtime,FCcount) + 2 * PartTime(FBtime,FBcount)) / 0.55 + PartTime(RB20time,RB20count);
                }
                else if (FCtime == 0 && FBtime == 0 && RBtime != 0 && RCtime != 0)
                {
                    return (2 * RBtime / RCcount + 2 * PartTime(RCtime,RCcount)) / 0.45 + PartTime(RB20time,RB20count);
                }
                else if (FCtime != 0 && FBtime == 0 && RBtime == 0 && RCtime == 0)
                {
                    return PartTime(FCtime,FCcount) / 0.1 + PartTime(RB20time,RB20count); 
                }
                else if (FCtime == 0 && FBtime != 0 && RBtime == 0 && RCtime == 0)
                {
                    return (PartTime(FBtime,FBcount)) / 0.17 + PartTime(RB20time,RB20count);
                }
                else if (FCtime == 0 && FBtime == 0 && RBtime != 0 && RCtime == 0)
                {
                    return (PartTime(RBtime,RBcount)) / 0.1 + PartTime(RB20time,RB20count);
                }
                else if (FCtime == 0 && FBtime == 0 && RBtime == 0 && RCtime != 0)
                {
                    return (PartTime(RCtime,RCcount)) / 0.1 + PartTime(RB20time,RB20count);
                }
                else if (FCtime != 0 && FBtime == 0 && RBtime != 0 && RCtime != 0)
                {
                    return ((2 * PartTime(RBtime,RBcount) + 2 * PartTime(RCtime,RCcount)) + 2 * PartTime(FCtime,FCcount)) / 0.66 + PartTime(RB20time,RB20count); ;
                }
                else if (FCtime == 0 && FBtime != 0 && RBtime != 0 && RCtime != 0)
                {
                    return ((2 * PartTime(RBtime,RBcount) + 2 * PartTime(FBtime,FBcount)) + 2 * (PartTime(FBtime,FBcount))) / 0.8 + PartTime(RB20time,RB20count);
                }
                else if (FCtime != 0 && FBtime != 0 && RBtime == 0 && RCtime != 0)
                {
                    return (2 * PartTime(FCtime,FCcount) + 2 * PartTime(FBtime,FBcount)) / 0.55 / 2 + (PartTime(RCtime,RCcount)) / 0.1 / 2 + PartTime(RB20time,RB20count);
                }
                else if (FCtime != 0 && FBtime != 0 && RBtime != 0 && RCtime == 0)
                {
                    return ((2 * PartTime(FCtime,FCcount) + 2 * PartTime(FBtime,FBcount)) / 0.55 / 2 + (PartTime(RBtime,RBcount)) / 0.1 / 2) + PartTime(RB20time,RB20count);
                }
                else if (FCtime == 0 && FBtime != 0 && RBtime == 0 && RCtime != 0)
                {
                    return (double)(((PartTime(FBtime,FBcount)) / 0.1 / 2 + (PartTime(RCtime,RCcount)) / 0.1) / 2) + PartTime(RB20time,RB20count);
                }
                else if (FCtime != 0 && FBtime == 0 && RBtime == 0 && RCtime != 0)
                {
                    return (double)((PartTime(FCtime,FCcount) / 0.17 / 2 + (PartTime(RCtime,RCcount)) / 0.1) / 2) + PartTime(RB20time,RB20count);
                }
                else if (FCtime == 0 && FBtime != 0 && RBtime != 0 && RCtime == 0)
                {
                    return (double)(((PartTime(FBtime,FBcount)) / 0.1 / 2 + (PartTime(RBtime,RBcount)) / 0.1) / 2) + PartTime(RB20time,RB20count);
                }
                else if (FCtime != 0 && FBtime == 0 && RBtime != 0 && RCtime == 0)
                {
                    return (double)((PartTime(FCtime,FCcount) / 0.17 / 2 + (PartTime(RBtime,RBcount)) / 0.1) / 2) + PartTime(RB20time,RB20count);
                }
                else
                {
                    return 0;
                }
            }else
            {
                if (FCtime != 0 && FBtime != 0 && RBtime != 0 && RCtime != 0)
                {
                    return 2 * (PartTime(FCtime,FCcount) + PartTime(FBtime,FBcount) + PartTime(RBtime,RBcount) + PartTime(RCtime,RCcount));
                }
                if (FCtime != 0 && FBtime != 0 && RBtime == 0 && RCtime == 0)
                {
                    return (2 * PartTime(FCtime,FCcount) + 2 * PartTime(FBtime,FBcount)) / 0.55;
                }
                else if (FCtime == 0 && FBtime == 0 && RBtime != 0 && RCtime != 0)
                {
                    return (2 * RBtime / RCcount + 2 * PartTime(RCtime,RCcount)) / 0.45;
                }
                else if (FCtime != 0 && FBtime == 0 && RBtime == 0 && RCtime == 0)
                {
                    return PartTime(FCtime,FCcount) / 0.1;
                }
                else if (FCtime == 0 && FBtime != 0 && RBtime == 0 && RCtime == 0)
                {
                    return (PartTime(FBtime,FBcount)) / 0.17;
                }
                else if (FCtime == 0 && FBtime == 0 && RBtime != 0 && RCtime == 0)
                {
                    return (PartTime(RBtime,RBcount)) / 0.1;
                }
                else if (FCtime == 0 && FBtime == 0 && RBtime == 0 && RCtime != 0)
                {
                    return (PartTime(RCtime,RCcount)) / 0.1;
                }
                else if (FCtime != 0 && FBtime == 0 && RBtime != 0 && RCtime != 0)
                {
                    return ((2 * PartTime(RBtime,RBcount) + 2 * PartTime(RCtime,RCcount)) + 2 * PartTime(FCtime,FCcount)) / 0.66;
                }
                else if (FCtime == 0 && FBtime != 0 && RBtime != 0 && RCtime != 0)
                {
                    return ((2 * PartTime(RBtime,RBcount) + 2 * PartTime(FBtime,FBcount)) + 2 * (PartTime(FBtime,FBcount))) / 0.8;
                }
                else if (FCtime != 0 && FBtime != 0 && RBtime == 0 && RCtime != 0)
                {
                    return (2 * PartTime(FCtime,FCcount) + 2 * PartTime(FBtime,FBcount)) / 0.55 / 2 + (PartTime(RCtime,RCcount)) / 0.1 / 2;
                }
                else if (FCtime != 0 && FBtime != 0 && RBtime != 0 && RCtime == 0)
                {
                    return ((2 * PartTime(FCtime,FCcount) + 2 * PartTime(FBtime,FBcount)) / 0.55 / 2 + (PartTime(RBtime,RBcount)) / 0.1 / 2);
                }
                else if (FCtime == 0 && FBtime != 0 && RBtime == 0 && RCtime != 0)
                {
                    return (double)(((PartTime(FBtime,FBcount)) / 0.1 / 2 + (PartTime(RCtime,RCcount)) / 0.1) / 2);
                }
                else if (FCtime != 0 && FBtime == 0 && RBtime == 0 && RCtime != 0)
                {
                    return (double)((PartTime(FCtime,FCcount) / 0.17 / 2 + (PartTime(RCtime,RCcount)) / 0.1) / 2);
                }
                else if (FCtime == 0 && FBtime != 0 && RBtime != 0 && RCtime == 0)
                {
                    return (double)(((PartTime(FBtime,FBcount)) / 0.1 / 2 + (PartTime(RBtime,RBcount)) / 0.1) / 2);
                }
                else if (FCtime != 0 && FBtime == 0 && RBtime != 0 && RCtime == 0)
                {
                    return (double)((PartTime(FCtime,FCcount) / 0.17 / 2 + (PartTime(RBtime,RBcount)) / 0.1) / 2);
                }
                else
                {
                    return 0;
                }

            }
           



        }
    }
}
