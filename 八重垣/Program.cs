/*
 2013/1/26
 一応形にはなりました。
 参照するファイルに無い番号の文字があるとそこだけ空白になるから注意が必要です。
    第XX期 コンピューター部 部長 XX
 */
/*
 2013/5/12
 最終的な確認と説明に加筆修正しました。
 * 第XX期 X部 3-X XX番 XX
 */
using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;//エクセルのネイムスペース

namespace 八重垣
{
    class Program
    {
        public class BASIC_information
        {
            public char[] SELL_SPOT_WORD;
            public int horizontal_number;//操作するエクセルの横のセル数を保存するメンバ変数
            public int vertical_number;//捜査するエクセルの縦のセル数を保存するメンバ変数


            //セルの横文字の選定
            public string SELL_WORD(int a)
            {
                int b = a / 27;
                int c = a % 27;
                if (c == 0) { c++; }//これを外すと大変なことになる。cが0の場合、SELL_SPOT_WORDの文字無しが連結されて x. y. z.A ,AA,AB…みたいになる。
                return String.Concat(SELL_SPOT_WORD[b], SELL_SPOT_WORD[c]);//文字列型を連結して返す
            }
            //縦列の数値と横列の英字をドッキングさせるメンバ関数
            public string Sell_SPOT(int a, string b)
            { return String.Concat(b, a.ToString()); }//文字列型を連結して返す

            //セルから抜き出した文字列に99,88が含んでいたならそこから六文字抜き出し
            public int include_check(string a)
            {
                int b, c = 0;
                if (a.IndexOf("99") >= 0) //990221
                {
                    b = a.IndexOf("99");
                    c = int.Parse(a.Substring(b, b + 6));
                }
                if (a.IndexOf("88") >= 0) //990221
                {
                    b = a.IndexOf("88");
                    c = int.Parse(a.Substring(b, b + 6));
                }
                return c;
            }

            //テキストファイルを指定して文字の取り出しとそれに付随する諸々の処理を司るメンバ関数
            public object Txt_read(int a, string b)
            {
                int c = a / 100;
                int d = a % 100;
                string e = b + c.ToString() + ".txt";//文字列型
                StreamReader txt_reader = new StreamReader(e, System.Text.Encoding.GetEncoding(932)/*文字コードを指定*/);//テキストファイルのオープン
                string Line = "";
                ArrayList arText = new ArrayList();
                int i = 1;
                while (Line != null)
                {
                    Line = txt_reader.ReadLine();//テキストファイルから読みだした一行を変数に保存
                    if (i == d) { break; }
                    i++;
                }
                txt_reader.Close();//開いたテキストファイルのクローズ
                object x = Line;
                Console.WriteLine(Line);
                return x;
            }


            //このクラスのコンストラクタ
            public BASIC_information()
            {
                SELL_SPOT_WORD = new char[27] { ' ', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' };
            }
        };

        //メイン関数
        static int Main(string[] args)
        {
            //BASIC_informationクラスのインスタンス
            BASIC_information BASIC = new BASIC_information();

            //操作に必要なデータの入力
            Console.WriteLine("最終更新1/26\n名簿掲載以外を示す数字が紛れると強制終了の危険性あり");
            Console.WriteLine("指定されたもの以外のコマンドを入れると予期せぬ動作を起こしますが仕様です\nファイル指定の際、\\は\\\\として下さい\n(例:ローカルディスク(C:)直下のエクセルというフォルダ内のtest.xlsxを指す場合\n    C:\\\\エクセル\\\\test.xlsx\n");
            Console.WriteLine("操作するエクセルファイルを指定してください");
            string filename = Console.ReadLine();
            Console.WriteLine("操作するエクセルシートの番号を指定してください");
            int sheet_snum = int.Parse(Console.ReadLine());
            Console.WriteLine("計測する横のセルの数を半角数字で入力してください。");
            BASIC.horizontal_number = int.Parse(Console.ReadLine());
            Console.WriteLine("計測する縦のセルの数を半角数字で入力してください。");
            BASIC.vertical_number = int.Parse(Console.ReadLine());
            Console.WriteLine("参照する名簿ファイルの場所を入力してください。");
            string file_point = Console.ReadLine();
            /*Console.WriteLine("ログを保存するテキストファイルの場所を入力して下さい");
            string log_name = Console.ReadLine();*/

            int vertical_sell = 0;//縦のセル番を入れる変数
            string horizontal_sell;//横のセル番号を入れる変数
            //Excel.application のインスタンスを作成
            Excel.Application App = new Excel.Application();
            string spot;

            if (App != null)
            {
                //Excelを表示
                App.Visible = true;
                //ワークブックのインスタンスを生成
                Excel.Workbook WB = App.Workbooks.Open(Filename: filename);
                //一番目のワークシートを選択?
                ((Excel.Worksheet)WB.Sheets[sheet_snum]).Select();
                Excel.Range rng;

                for (int n = 1; n <= BASIC.horizontal_number; n++)
                {
                    //横の文字を入れる
                    horizontal_sell = BASIC.SELL_WORD(n);
                    for (int i = 1; i <= BASIC.vertical_number; i++)
                    {
                        //縦の数字を入れる
                        vertical_sell = i;
                        //ここで文字と番号を連結
                        spot = BASIC.Sell_SPOT(vertical_sell, horizontal_sell);
                        rng = App.get_Range(spot);//セルから取り出した文字列を変数に代入

                        if (rng.Value2 != null)//rngの中身がnull以外なら以下の処理
                        {
                            string text_ = rng.Value2.ToString();//rng.Value2はオブジェクト型
                            int trans = BASIC.include_check(text_);//抜き出した数字を代入。評価する数字以外なら0を返す
                            if (trans != 0)
                            {
                                rng.Value2 = BASIC.Txt_read(trans, file_point);//テキストに関するメンバ関数を呼び出して返り値をセルに書き込む
                            }
                            Console.WriteLine(text_);//読み込んだ文字をコンソールに表示
                        }
                    }
                }
                WB.Close(true);//エクセルワークブックのクローズ
            }
            //エクセルの終了
            App.Quit();
            //Appオブジェクトの破棄
            System.Runtime.InteropServices.Marshal.ReleaseComObject(App);
            return 0;
        }
    }
}