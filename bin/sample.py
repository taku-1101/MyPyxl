import MyPyxl
import pandas as pd


#データ読み込み
df_afgr = pd.read_csv(r'.\in\aftergrow.csv', engine='python')
df_hlhp = pd.read_csv(r'.\in\hellohappy.csv', engine='python')

#テンプレートの読み込み
pyxl = MyPyxl.MyPyxl(r'.\in\tmplete.xlsx')

for d in [df_afgr, df_hlhp]:

        #書き込む情報を分ける
        list_name = d[['名前','パート']].values.tolist()
        list_pro = d[['誕生日','星座','身長']].values.tolist()
        list_com = d['コメント'].values.tolist()
        list_img = d['画像'].values.tolist()

        #Excelで作成した名前の定義の取得
        mytset = pyxl.get_Defines()

        for i in range(5):
                #水平方向にずらしながらセルに書き込む
                mytset['LOOPR4_ROW_name'].append(list_name[i])
                #垂直方向にずらしながらセルに書き込む
                mytset['LOOPR5_COL_pro'].append(list_pro[i])
                #1つのセルに書き込む
                mytset['LOOPR5_CELL_comment'].append(list_com[i])
                #指定のセルに画像を貼り付ける
                mytset['LOOPR2_IMG_chara'].append(list_img[i])

        #作成前のおまじない
        pyxl.regist_Defines2Range('RNGD2_member', mytset)

#エクセル作成
pyxl.create_xlsx(['RNGD2_member'],r'.\output.xlsx')


