import streamlit as st
from pulp import *
from openpyxl import *
import pandas as pd
import time

st.title('ライブ裏方シフトのシフト表')
st.header('▼ 欠席者の入力')

#ダウンロードしたファイルの読み込み
uploaded_file = st.file_uploader("シフト表、入力データをアップロード", type='xlsx')
if uploaded_file is not None:
    data_nyuuryoku = pd.read_excel(uploaded_file, sheet_name=0)
    data_shift = pd.read_excel(uploaded_file, sheet_name=1, index_col=0)
    st.table(data_nyuuryoku)
    st.table(data_shift)

#ファイルから情報を取得
#メンバーを取得
    st.write('メンバーの名前')
    member_list = []
    for i in range(0, len(data_nyuuryoku)):
        if member_list.count(data_nyuuryoku.iloc[i, 0]) == 0:
            member_list.append(data_nyuuryoku.iloc[i, 0])
    st.table(member_list)

#バンド数を取得
    st.write('バンド数')
    st.write(len(data_shift.columns))
        
#欠席者の名前を入力
    with st.form('input_absentee_name', clear_on_submit=True):
        st.write('メンバーの名前を入力してください')
        member_name = st.text_input('↓↓↓')
        submitted = st.form_submit_button('決定')
        if submitted:
            data_nyuuryoku_2 = data_nyuuryoku.drop(member_list.index(member_name))
            member_list.remove(member_name)
        st.table(member_list)
        st.table(data_nyuuryoku_2)
        
    re_member = len(member_list)
    re_slot = len(data_shift.columns)

#定数用のデータの作成(欠席者対応版)
    I = [i+1 for i in range(re_member)]
    J = [i+1 for i in range(4)]
    T = [i+1 for i in range(re_slot)]
    st.write(len(I))
    st.write(len(T))
    
    data_nyuuryoku_3 = data_nyuuryoku_2.drop(columns = '名前')
    st.table(data_nyuuryoku_3)
    a = {}    
    for i in range(0, len(I)):
        for t in range(0, len(T)):
            a[i, t] = data_nyuuryoku_2.iloc[i, t+1]
    st.write(a[0, 7])
    
    old_x = {}
    for j in range(0, len(J)):
        for t in range(0, len(T)):
            old_x[j, t] = data_shift.iloc[j, t]
    #st.write(old_x[3, 7])
            
#空問題の作成
    model = LpProblem('Shift', sense=LpMinimize)

#決定変数の作成
    x = {}
    for i in I:
        for j in J:
            for t in T:
                x[i, j, t] = LpVariable(f'x{i},{j},{t}', cat=LpBinary)
    v = {}
    for i in I:
        v[i] = LpVariable(f'v{i}', lowBound=0)
    w = {}
    for i in I:
        for t in T:
            w[i, t] = LpVariable(f'w{i}, {t}', cat=LpBinary)
    z = {}
    for i in I:
        for t in T:
            z[i, t] = LpVariable(f'z{i}, {t}', cat=LpBinary)
            
#制約条件の追加
#条件①
    for j in J:
        for t in T:
            model += lpSum(x[i, j, t] for i in I) == 1
#条件②
    for i in I:
        for t in T:
            model += lpSum(x[i, j, t] for j in J) <= 1
#条件③
    for i in (0, len(I)-1):
        for t in (0, len(T)-1):
            if a[i, t] == 0:
                model += lpSum(x[i, j, t] for j in J) == 0
#条件④
    for i in (0, len(I)-1):
        for t in (0, len(T)-1):
            if a[i, t] == 3:
                model += lpSum(x[i, j, t] for j in J) == 0
#条件⑤
    s = (len(J) * len(T)) / len(I)
    for i in I:
        model += lpSum(x[i, j, t] for j in J for t in T) >= s - v[i]
        model += lpSum(x[i, j, t] for j in J for t in T) <= s + v[i]
#条件⑥
    for i in I:
        for t in range(len(T)-1):
            model += lpSum(x[i, j, t+1] for j in J) + lpSum(x[i, j, t+2] for j in J) <= 1 + z[i, t+1]
#条件⑦
    for i in (0, len(I)-1):
        for t in (0, len(T)):
            if a[i, t] == 2:
                model += lpSum(x[i, j, t] for j in J) == 0 + w[i, t]
#条件⑧
    for i in I:
        for j in J:
            for t in T:
                model += lpSum(x[i, j, t])
        
#目的関数の設定
    model += lpSum(v[i] for i in I) + lpSum(z[i, t] for i in I for t in T) + lpSum(w[i, t] for i in I for t in T)

#最適化の実行
    model.solve()