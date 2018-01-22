using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ILOG.Concert;
using ILOG.CPLEX;
using System.IO;
using System.Data;
using System.Data.OleDb;

namespace RLND_1
{

    class Program
    {
        #region
        // Indices and sets
        public const int customer = 20;         //customer i
        public const int cpoint = 8;            //collection point j
        public const int rcenter = 5;          //recycling center k
        public const int ptype = 3;             //product type p
        public const int scenario = 1;         //scenario s
        public const int period = 1;            //time period t
        public const int elevel = 3;			//environment protection level option for a recycling center c
        public const int capacity = 3;          //capacity level option for a recycling center g
        public const int vlength = 7;          //length of a collection period

        //Parameters
        // public double[] a = new double[cpoint];     //annual renting cost of collection point j
        public double[] a = new double[cpoint];
        
        // public double[] b = new double[ptype];      //unit daily inventory carrying cost of product p
        public double[] b = new double[ptype];      //unit daily inventory carrying cost of product p

        // public double[] u = new double[ptype];      //unit collection cost of product p
        public double[] u = new double[ptype];      //unit collection cost of product p

        // public int[,,,] r = new int[customer, ptype, period, scenario];       //daily volume of product p returned by customer i during time period t in scenario s
        public int[,,,] r = new int[customer, ptype, period, scenario] ;

        public const int wday = 420;                //annual working days per time period a-j

        // public double[,,] z = new double[cpoint, rcenter, ptype];                   //unit transportation cost of product p from collection point j to recycling center k
        public double[,,] z = new double[cpoint, rcenter, ptype];

        // public double[] h = new double[ptype];                                   //unit handling cost of product p
        public double[] h = new double[ptype];


        // public double[,,] q = new double[rcenter, elevel, capacity];                //daily volume of product p returned by customer i during time period t in scenario s
        public double[,,] q = new double[rcenter, elevel, capacity];

        // public double[,] d1 = new double[customer, cpoint];                     //distance from customer i to collection point j
        public double[,] d1 = new double[customer, cpoint];

        public double[,] d2 = new double[cpoint, rcenter];                      //distance from collection point j to recycling center k

        public const int l = 25;                //maximum allowed distance from a given customer to a collection point

        // public double[] e1 = new double[cpoint];                                //unit daily carbon emission (in kg) from operating collection point j
        public double[] e1 = new double[cpoint];                                //unit daily carbon emission (in kg) from operating collection point j


        public double[,,] e2 = new double[ptype, cpoint, rcenter];              //unit carbon emission (in kg) from transporting product p collection point j to recycling center k


        public double[,] e3 = new double[rcenter, elevel];                      //unit daily carbon emission (in kg) from operating recycling center k with environment protection level c

        // public double[] e4 = new double[ptype];                                 //unit carbon emission (in kg) from handling product p  
        //public double[] e4 = new double[ptype] {0.8,1 };

        public const double e = 4000000000;                //allowed CO2 emission (in kg) for global reverse logistics network during a time period
        public const double lambda = 0.6;           //A coefficient lambda is introduced to combine these two objectives to balance cost and carbon emission.
        public const double delta = 0.4;
        #endregion

        #region
        public string strConnection;  //  access database connection
        public void input()
        {
            // Random rnd = new Random();

            strConnection = "Provider=Microsoft.Ace.OleDb.12.0;";
            strConnection += @"Data Source=E:\cplex_codes\RLND-1\datasets\RLND1.accdb";//这里用的是绝对路径
            OleDbConnection objConnection = new OleDbConnection(strConnection);
            objConnection.Open();

            string str1 = "select * from r";
            OleDbCommand myCommand001 = new OleDbCommand(str1, objConnection);
            OleDbDataReader myReader001 = myCommand001.ExecuteReader();
            int d11 = 0;
            int d12 = 0;
            int d13 = 0;
            int d14 = 0;
            while (myReader001.Read())
            {
                d11 = myReader001.GetInt32(1) - 1;
                d12 = myReader001.GetInt32(2) - 1;
                d13 = myReader001.GetInt32(3) - 1;
                d14 = myReader001.GetInt32(4) - 1;
                r[d11, d12, d13, d14] = myReader001.GetInt32(5);
            }
            myReader001.Close();

            string str2 = "select * from z";
            OleDbCommand myCommand002 = new OleDbCommand(str2, objConnection);
            OleDbDataReader myReader002 = myCommand002.ExecuteReader();
            int d15 = 0;
            int d16 = 0;
            int d17 = 0;
            while (myReader002.Read())
            {
                d15 = myReader002.GetInt32(1) - 1;
                d16 = myReader002.GetInt32(2) - 1;
                d17 = myReader002.GetInt32(3) - 1;
                z[d15,d16,d17] = myReader002.GetInt32(4);
            }
            myReader002.Close();

            string str3 = "select * from q";
            OleDbCommand myCommand003 = new OleDbCommand(str3, objConnection);
            OleDbDataReader myReader003 = myCommand003.ExecuteReader();
            int d111 = 0;
            int d112 = 0;
            int d113 = 0;
            while (myReader003.Read())
            {
                d111 = myReader003.GetInt32(1) - 1;
                d112 = myReader003.GetInt32(2) - 1;
                d113 = myReader003.GetInt32(3) - 1;
                q[d111, d112, d113] = myReader003.GetInt32(4);
            }
            myReader003.Close();

            string str4 = "select * from d1";
            OleDbCommand myCommand004 = new OleDbCommand(str4, objConnection);
            OleDbDataReader myReader004 = myCommand004.ExecuteReader();
            int d21 = 0;
            int d22 = 0;
            while (myReader004.Read())
            {
                d21 = myReader004.GetInt32(1) - 1;
                d22 = myReader004.GetInt32(2) - 1;
                d1[d21, d22] = myReader004.GetInt32(3);
            }
            myReader004.Close();

            string str5 = "select * from d2";
            OleDbCommand myCommand005 = new OleDbCommand(str5, objConnection);
            OleDbDataReader myReader005 = myCommand005.ExecuteReader();
            int d31 = 0;
            int d32 = 0;
            while (myReader005.Read())
            {
                d31 = myReader005.GetInt32(1) - 1;
                d32 = myReader005.GetInt32(2) - 1;
                d2[d31, d32] = myReader005.GetInt32(3);
            }
            myReader005.Close();

            string str6 = "select * from e2";
            OleDbCommand myCommand006 = new OleDbCommand(str6, objConnection);
            OleDbDataReader myReader006 = myCommand006.ExecuteReader();
            int d41 = 0;
            int d42 = 0;
            int d43 = 0;
            while (myReader006.Read())
            {
                d41 = myReader006.GetInt32(1) - 1;
                d42 = myReader006.GetInt32(2) - 1;
                d43 = myReader006.GetInt32(3) - 1;
                e2[d41, d42 ,d43] = myReader006.GetInt32(4);
            }
            myReader006.Close();

            string str7 = "select * from e3";
            OleDbCommand myCommand007 = new OleDbCommand(str7, objConnection);
            OleDbDataReader myReader007 = myCommand007.ExecuteReader();
            int d51 = 0;
            int d52 = 0;
            while (myReader007.Read())
            {
                d51 = myReader007.GetInt32(1) - 1;
                d52 = myReader007.GetInt32(2) - 1;
                e3[d51, d52] = myReader007.GetInt32(3);
            }
            myReader007.Close();



            objConnection.Close();
        }
        #endregion

        void solve()
        {
            Cplex Model = new Cplex();

            // decision variables
            #region
            //x1-tjs (first stage t = 1),length of a collection period (in days) at collection point j during time period t
            //INumVar[][][] X1 = new INumVar[period][][];
            //for (int i = 0; i < period; i++)
            //{
            //    X1[i] = new INumVar[cpoint][];
            //    for (int ii = 0; ii < cpoint; ii++)
            //    {
            //        X1[i][ii] = new INumVar[scenario];
            //        X1[i][ii] = Model.NumVarArray(scenario, 1, 7, NumVarType.Int);
            //    }
            //}

            //x2-tpjks ,volume of products returned from collection point j to recycling center k during time period t
            INumVar[][][][][] X2 = new INumVar[period][][][][];
            for (int a = 0; a < period; a++)
            {
                X2[a] = new INumVar[ptype][][][];
                for (int aa = 0; aa < ptype; aa ++)
                {
                    X2[a][aa] = new INumVar[cpoint][][];
                    for (int bb = 0; bb < cpoint; bb++)
                    {
                        X2[a][aa][bb] = new INumVar[rcenter][];
                        for (int cc = 0; cc < rcenter; cc++)
                        {
                            X2[a][aa][bb][cc] = new INumVar[scenario];
                            X2[a][aa][bb][cc] = Model.NumVarArray(scenario, 0, System.Double.MaxValue, NumVarType.Int);
                         }
                     }
                }             
            }

            //x3-tij ,Binary decision variable equals ‘1’ if customer i is allocated to collection point j during time period t and ‘0’ otherwise
            INumVar[][][] X3 = new INumVar[period][][];
            for (int aa = 0; aa < period; aa++)
            {
                X3[aa] = new INumVar[customer][];
                for (int bb = 0; bb < customer; bb++)
                {
                    X3[aa][bb] = new INumVar[cpoint];
                    X3[aa][bb] = Model.NumVarArray(cpoint, 0.0, 1.0, NumVarType.Bool);
                }
            }


            //x4-tj ,Binary decision variable equals ‘1’ if collection point j is rented during time period t and ‘0’ otherwise
            INumVar[][] X4 = new INumVar[period][];
            for (int i = 0; i < period; i++)
            {
                X4[i] = new INumVar[cpoint];
                X4[i] = Model.NumVarArray(cpoint, 0.0, 1.0, NumVarType.Bool);
            }

            //x5-k ,Binary decision variable equals ‘1’ if recycling center k is established and ‘0’ otherwise
            INumVar[] X5 = new INumVar[rcenter];
            X5 = Model.NumVarArray(rcenter, 0, 1, NumVarType.Bool);

            //x6 maximum capacity level option for recycling center k
            INumVar[] X6 = new INumVar[rcenter];
            X6 = Model.NumVarArray(rcenter, 1, 3, NumVarType.Int);

            //X7-tjs Auxiliary variable  x7 the frequency of recylce during a period
            //INumVar[][][] X7 = new INumVar[period][][];
            //for (int i = 0; i < period; i++)
            //{
            //    X7[i] = new INumVar[cpoint][];
            //    for (int ii = 0; ii < cpoint; ii++)
            //    {
            //        X7[i][ii] = new INumVar[scenario];
            //        X7[i][ii] = Model.NumVarArray(scenario, 30, wday, NumVarType.Int);
            //    }
            //}

            #endregion
           
            Double M = 100000;

            // constraints
            #region
            // formulation （4）
            INumExpr[] expr1 = new INumExpr[1];
            
            for (int aa = 0; aa < period; aa++)
            {
                for (int bb = 0; bb < customer; bb++)
                {
                    expr1[0] = X3[0][0][0];
                    for (int cc = 0; cc < cpoint; cc++)
                    {
                        expr1[0] = Model.Sum(expr1[0], X3[aa][bb][cc]);
                    }
                    expr1[0] = Model.Sum(expr1[0], Model.Prod(-1.0, X3[0][0][0]));
                    Model.AddEq(expr1[0], 1);
                }
            }


            // formulation （5）
            INumExpr[] expr2 = new INumExpr[1];
            INumExpr[] expr3 = new INumExpr[1];
            for (int aa = 0; aa < period; aa++)
            {

                for (int bb = 0; bb < cpoint; bb++)
                {
                    expr2[0] = X3[0][0][0];
                    expr3[0] = X4[0][0];
                    for (int cc = 0; cc < customer; cc++)
                    {
                        Model.Sum(expr2[0], X3[aa][cc][bb]);
                    }
                    expr2[0] = Model.Sum(expr2[0], Model.Prod(-1.0, X3[0][0][0]));
                    expr3[0] = Model.Prod(M, X4[aa][bb]);
                    expr3[0] = Model.Sum(expr3[0], Model.Prod(-1.0, X4[0][0]));
                    Model.AddLe(expr2[0], expr3[0]);
                }

            }

            // formulation （6）
            INumExpr[] expr4 = new INumExpr[1];
            INumExpr[] expr5 = new INumExpr[1];
            INumExpr[] expr6 = new INumExpr[1];
            for (int aa = 0; aa < period; aa++)
            {
                for (int cc = 0; cc < scenario; cc++)
                {
                    for (int bb = 0; bb < cpoint; bb++)
                    {
                        expr5[0] = X3[0][0][0];
                        for (int dd = 0; dd < customer; dd++)
                        {
                            for (int ee = 0; ee < ptype; ee++)
                            {
                                expr5[0] = Model.Sum(expr5[0], Model.Prod(Model.Prod(r[dd, ee, aa, cc], X3[aa][dd][bb]), vlength));
                            }
                        }
                        expr5[0] = Model.Sum(expr5[0], Model.Prod(-1.0, X3[0][0][0]));

                        expr6[0] = X2[0][0][0][0][0];
                        for (int ff = 0; ff < rcenter; ff++)
                        {
                            for (int ee = 0; ee < ptype; ee++)
                            {
                                expr6[0] = Model.Sum(expr6[0], X2[aa][ee][bb][ff][cc]);
                            }
                        }
                        expr6[0] = Model.Sum(expr6[0], Model.Prod(-1.0, X2[0][0][0][0][0]));
                        Model.AddEq(expr5[0], expr6[0]);
                    }

                }
            }

            // formulation （7-1）
            INumExpr[] expr7 = new INumExpr[1];
            for (int aa = 0; aa < period; aa++)
            {
                for (int bb = 0; bb < rcenter; bb++)
                {
                    for (int cc = 0; cc < scenario; cc++)
                    {
                        for (int dd = 0; dd < cpoint; dd++)
                        {
                            expr7[0] = X2[0][0][0][0][0];
                            for (int ee = 0; ee < ptype; ee++)
                            {
                                expr7[0] = Model.Sum(expr7[0], X2[aa][ee][dd][bb][cc]);
                            }
                            expr7[0] = Model.Sum(expr7[0], Model.Prod(-1.0, X2[0][0][0][0][0]));
                            Model.AddLe(expr7[0], Model.Prod(X5[bb], M));
                        }
                    }
                      

                }
            }

            // formulation （7-2）
            INumExpr[] expr71 = new INumExpr[1];
            for (int aa = 0; aa < period; aa++)
            {
                for (int bb = 0; bb < rcenter; bb++)
                {
                    for (int cc = 0; cc < scenario; cc++)
                    {
                        
                        for (int dd = 0; dd < cpoint; dd++)
                        {
                            expr71[0] = X2[0][0][0][0][0];
                            for (int ee = 0; ee < ptype; ee++)
                            {
                                expr71[0] = Model.Sum(expr71[0], X2[aa][ee][dd][bb][cc]);
                            }
                            expr71[0] = Model.Sum(expr71[0], Model.Prod(-1.0, X2[0][0][0][0][0]));
                            Model.AddLe(expr71[0],Model.Sum(Model.Prod(X6[bb],1000), Model.Prod(Model.Sum(1,Model.Prod(X5[bb], -1)),M)));
                        }  
                    }
                }
            }

            // formulation （8）
            INumExpr[] expr8 = new INumExpr[1];
            for (int a = 0; a < period; a++)
            {
                expr8[0] = X4[0][0];
                for (int b = 0; b < cpoint; b++)
                {
                    expr8[0] = Model.Sum(expr8[0], X4[a][b]);
                }
                expr8[0] = Model.Sum(expr8[0], Model.Prod(-1.0, X4[0][0]));
                Model.AddGe(expr8[0], 1);
            }

            // formulation （9）
            INumExpr[] expr9 = new INumExpr[1];
            expr9[0] = X5[0];
            for (int a = 0; a < rcenter; a++)
            {
                expr9[0] = Model.Sum(expr9[0], X5[a]);
            }
            expr9[0] = Model.Sum(expr9[0], Model.Prod(-1.0, X5[0]));
            Model.AddGe(expr9[0], 1);

            // formulation （10）
            INumExpr[] expr10 = new INumExpr[1];
            for (int a = 0; a < period; a++)
            {
                for (int b = 0; b < rcenter; b++)
                {
                    for (int c = 0; c < scenario; c++)
                    {
                        
                        for (int d = 0; d < cpoint; d++)
                        {
                            expr10[0] = X2[0][0][0][0][0];
                            for (int e = 0; e < ptype; e++)
                            {
                                expr10[0] = Model.Sum(expr10[0], X2[a][e][d][b][c]);
                            }
                            expr10[0] = Model.Sum(expr10[0], Model.Prod(-1.0, X2[0][0][0][0][0]));
                            Model.AddGe(expr10[0], X5[b]);  
                        } 
                    }
                }
            }


            // formulation （11）
            for (int a = 0; a < period; a++)
            {
                for (int c = 0; c < customer; c++)
                {
                    for (int b = 0; b < cpoint; b++)
                    {
                        Model.AddLe(Model.Prod(d1[c, b], X3[a][c][b]), l);
                    }
                }
            }

            // formulation （12）
            for (int a = 0; a < period; a++)
            {
                for (int aa = 0; aa < ptype; aa++)
                {
                   for (int b = 0; b < cpoint; b++)
                   {
                        for (int c = 0; c < rcenter; c++)
                        {
                            for (int d = 0; d < scenario; d++)
                            {
                                Model.AddGe(X2[a][aa][b][c][d], 0);
                            }
                        }
                   } 
                }
            }


            
            // #endregion
            //formulation (15)  //formulation (1)  objective function -1
            // collection points rent cost
           
            INumExpr[] expr11 = new INumExpr[1];
            for (int a = 0; a < period; a++)
            {
                expr11[0] = X4[0][0];
                for (int b = 0; b < cpoint; b++)
                {
                    expr11[0] = Model.Sum(expr11[0], Model.Prod(X4[a][b], e1[b]));
                }
                expr11[0] = Model.Sum(expr11[0], Model.Prod(-1.0, X4[0][0]));
            }

            INumExpr[] expr12 = new INumExpr[1];
            INumExpr[] expr121 = new INumExpr[1];

            for (int a = 0; a < period; a++)
            {
               
                for (int b = 0; b < rcenter; b++)
                {
                    expr121[0] = X5[0];
                    for (int c = 0; c < cpoint; c++)
                    {
                        for (int d = 0; d < ptype; d++)
                        {
                            for (int e = 0; e < customer; e++)
                            {
                                expr12[0] = X2[0][0][0][0][0];
                                for (int f = 0; f < scenario; f++)
                                {
                                    expr12[0] = Model.Sum(expr12[0], Model.Prod(Model.Prod(X2[a][d][c][b][f], e2[d, c, b]), wday/vlength));
                                }
                                expr12[0] = Model.Sum(expr12[0], Model.Prod(-1.0, X2[0][0][0][0][0]));
                            }
                        }
                    }expr121[0] = Model.Sum(expr121[0], Model.Prod(expr12[0], X5[b]));
                    
                }
            }

            INumExpr[] expr13 = new INumExpr[1];
            for (int a = 0; a < period; a++)
            {

                for (int b = 0; b < rcenter; b++)
                {
                    expr13[0] = X5[0];
                    for (int bb = 0; bb < elevel; bb++)
                    {
                        expr13[0] = Model.Sum(expr13[0], Model.Prod(X5[b], e3[b, bb]));
                    }
                    expr13[0] = Model.Sum(expr13[0], Model.Prod(-1.0, X5[0]));
                }
            }

  
              Model.AddLe(Model.Prod(wday, Model.Sum(expr11[0], expr13[0])), e);//expr12[0],  ,)
            // #endregion

            // //formulation (1)  objective function -1
            // #region
            INumExpr[] expr15 = new INumExpr[1];
            expr15[0] = X4[0][0];
            for (int i = 0; i < period; i++)
            {
                for (int ii = 0; ii < cpoint; ii++)
                {
                    expr15[0] = Model.Sum(expr15[0], Model.Prod(a[ii], X4[i][ii]));
                } 
            }
            expr15[0] = Model.Sum(expr15[0], Model.Prod(-1.0, X4[0][0]));
            // ok

           
            INumExpr[] expr17 = new INumExpr[1];
           for (int a = 0; a < period; a++)
            {
                for (int aaa = 0; aaa < scenario; aaa++)
                {

                    for (int bb = 0; bb < cpoint; bb++)
                    {
                        for (int bbb = 0; bbb < customer; bbb++)
                        {
                            expr17[0] = X3[0][0][0];
                            for (int aa = 0; aa < ptype; aa++)
                            {
                              
                                expr17[0] = Model.Sum(expr17[0],Model.Prod(Model.Prod(r[bbb, aa, a, aaa] * b[aa], X3[a][bbb][bb]),(vlength+1)* 0.5));

                            }
                            expr17[0] = Model.Sum(expr17[0], Model.Prod(-1.0, X3[0][0][0]));
                        }
                    }
                }
            }

            INumExpr[] expr18 = new INumExpr[1];
            INumExpr[] expr41 = new INumExpr[1];
            expr18[0] = X5[0];
            for (int a = 0; a < period; a++)
            {
                
                for (int aa = 0; aa < rcenter; aa++)
                {  
                    for (int bb = 0; bb < cpoint; bb++)
                    {
                        for (int b = 0; b < ptype; b++)
                        {
                            for (int aaa = 0; aaa < customer; aaa++)
                            {
                                expr41[0] = X2[0][0][0][0][0];
                                for (int bbb = 0; bbb < scenario; bbb++)
                                {
                                    expr41[0] = Model.Sum(expr41[0], Model.Prod(Model.Prod(X2[a][b][bb][aa][bbb], z[bb, aa, b]),vlength));
                                }
                                expr41[0] = Model.Sum(expr17[0], Model.Prod(-1.0, X2[0][0][0][0][0]));
                            }
                        }
                    }
                    expr18[0] = Model.Prod(expr41[0], X5[aa]);  
                } 
            }
            expr18[0] = Model.Sum(expr18[0], Model.Prod(-1.0, X5[0]));

            INumExpr[] expr19 = new INumExpr[1];
            for (int i = 0; i < period; i++)
            {
                for (int ii = 0; ii < rcenter; ii++)
                {
                    for (int iii = 0; iii < elevel; iii++)
                    {
                        expr19[0] = X5[0];
                        for (int iiii = 0; iiii < capacity; iiii++)
                        {
                            expr19[0] = Model.Sum(expr19[0], Model.Prod(q[ii, iii, iiii], X5[ii]));
                        }
                        
                    }
                }
            }
            expr19[0] = Model.Sum(expr19[0], Model.Prod(-1.0, X5[0]));

       

            #endregion
            INumExpr[] expr21 = new INumExpr[1];
            expr21[0] = Model.Prod(Model.Sum(expr15[0], Model.Prod(expr17[0], wday),  expr18[0],expr19[0]),lambda); // 

            INumExpr[] expr22 = new INumExpr[1];
            expr22[0] = Model.Prod(Model.Sum(expr11[0],expr12[0],expr13[0]),delta);// Model.Prod(Model.Prod(wday, Model.Sum(expr11[0], expr12[0], expr13[0])),delta);

            Model.AddMinimize(Model.Sum(expr21[0],expr22[0]));//

            Model.ExportModel("RLND.LP");

             if (Model.Solve())
            {
                Console.WriteLine("statue=" + Model.GetStatus());
                Console.WriteLine("obj=" + Model.ObjValue);
                Console.WriteLine("X2的结果");

                for (int aa2 = 0; aa2 < period; aa2++)
                {
                    for (int bb2 = 0; bb2 < ptype; bb2++)
                    {
                        for (int cc = 0; cc < cpoint; cc++)
                        {
                            for (int dd = 0; dd < rcenter; dd++)
                            {
                                for (int ee = 0; ee < scenario; ee++)
                                {
                                    Console.WriteLine(Model.GetValue(X2[aa2][bb2][cc][dd][ee]));
                                    Console.WriteLine();
                                }
                                
                            }
                        }
                    }
                }

                //for (int a = 0; a < X6.Length; a++)
                //{
                //    Console.WriteLine("result[" + a + "] = " + Model.GetValue(X6[a]));
                //}
                //Console.WriteLine();


                Model.End();
            }
            else
            {
                Model.End();
                Console.WriteLine();
                Console.WriteLine("cannot be solved");
            }
        }



        static void Main(string[] args)
        {
            DateTime A = DateTime.Now;
            Console.WriteLine("CPLEX");

            Program prg = new Program();
            prg.input();
            prg.solve();

            DateTime B = DateTime.Now;
            System.TimeSpan CC = B - A;
            System.Console.WriteLine("RUNNING TIME :    {0}", CC);
            Console.Read();
        }
    }
}
