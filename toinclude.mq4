//+------------------------------------------------------------------+
//|                                                          Guu.mq4 |
//|                                            Copyright 2010 lab 42 |
//|                                                    http://l42.us |
//+------------------------------------------------------------------+
#property copyright "Copyright 2010 lab 42"
#property link      "http://l42.us"

#import "test2.dll"
   string GetIdeNumber();
   int SoftSX(string path,string a0, int a1, string a2, string a3, int a4, string a5, string a6, double a7);
#import

datetime d1;
int everytime = 10;

void init_time()
{
   d1 = TimeCurrent();
   
}
int check_reg()
{
   int gi_328;
   string gs_356="";
   string gs_320="";
   string gs_332=""; 
   bool gi_352;

   
   
   gs_320 = GetIdeNumber();
   if (!IsTesting()) {
      if (IsDemo()) gi_352 = FALSE;
      else gi_352 = TRUE;
      gi_328 = SoftSX(TerminalPath(), gs_320, 1, AccountName(), AccountNumber(), gi_352, AccountCompany(), AccountServer(), AccountBalance());
      if (gi_328 == -4) {
         Alert("智能交易系统尚未注册，请您安装注册！");
         return (-1);
      }
      if (gi_328 == -3) {
         Alert("软件使用权尚未开通，请与服务商联系！");
         return (-1);
      }
      if (gi_328 > 0 && gi_328 <= 5) Alert("软件使用期还有" + gi_328 + "天，请与服务商联系！");
      else {
         if (gi_328 == -1) Alert("软件使用期限已到，服务将随时停止！");
         else {
            if (gi_328 == 0) {
               Alert("连接远程服务器失败，请重新启动软件！");
               gs_356 = "ok";
            }
         }
      }
      return (1);
   }
}

int check_time() // check by every 2 Hour
{
   datetime dn;
   int res;
   res =0;
   dn = TimeCurrent();
   if( dn - d1 >= everytime )
   {
      d1 = dn;
      //check now
     
       res = check_reg(); return (res);
      /*
         return res;
      */
      
   }
   return (1);
}

