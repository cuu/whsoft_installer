#include "toinclude.mq4"

extern double �������� = 40.0;
extern double �Ե����� = 90.0;
extern double �������� = 15.0;
extern double �������� = 2.0;
extern double �൥���� = 0.1;
extern double �յ����� = 0.1;
extern int ��󶩵��� = 6;
extern double �Զ����� = 20.0;
extern int ���������� = 4;
extern double ƽ̨������С����=0.1;
extern int ֹͣ�µ� = 1;
extern int ǿ��ƽ�� = 1;


double b_buy = 88888.0;
double s_sell = 99999.0;
double g_ord_open_price_204;
double g_ord_open_price_212;
double g_lots_220;
double g_lots_228;
double g_ord_lots_236;
double g_ord_lots_244;
double g_price_252;
double gd_260;
double g_pos_268;
double g_ticket_300;
double g_ticket_308;
double gd_316;
double gd_324;
double gd_332;
double gd_356;
int gi_396;
int gi_400;

double bar_new;
double bar_old;

int buy_sell_cut=0;

string m_diskid;
int m_softzt;
string oldTime="";
string nowTime="";
int TimeSub=0;
int m_isDemo;

int start()
{
   
   if( check_time() != 1) return (-1); // check ���ɹ�
   
   return(0);
}
int init()
{
   if (!IsExpertEnabled()) Alert("���! û��*���ܽ���* ");
   init_time();
}
int deinit()
{

   return(0);
}