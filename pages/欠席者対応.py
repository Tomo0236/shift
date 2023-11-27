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
    
    st.write('出演情報データ')
    st.table(data_nyuuryoku)
    data_nyuuryoku = data_nyuuryoku.drop(columns = '名前')
    #st.table(data_nyuuryoku)
    st.write('前回作成したシフト表')
    st.table(data_shift)
    
    @st.cache_resource
    def data_nyuuryoku_cache():
        df = data_nyuuryoku.drop(member_list.index(member_name))
        return df
    
    @st.cache_resource
    def delete_member_list():
        return member_list
        
#欠席者の名前を入力
    member_list = delete_member_list()
    with st.form('input_absentee_name', clear_on_submit=True):
        st.write('メンバーの名前を入力してください')
        member_name = st.text_input('↓↓↓')
        submitted = st.form_submit_button('決定')
        data_nyuuryoku = data_nyuuryoku_cache()
        if submitted:
            member_list.remove(member_name)
        st.table(member_list)
        st.table(data_nyuuryoku)
        
    re_member = len(member_list)
    re_slot = len(data_shift.columns)
    
    def all_clear():
        delete_member_list.clear()
        data_nyuuryoku_cache.clear()
    
    btn = st.button('clear turn', on_click=all_clear)

#定数用のデータの作成(欠席者対応版)
    I = [i+1 for i in range(re_member)]
    J = [i+1 for i in range(4)]
    T = [i+1 for i in range(re_slot)]
    #st.write(range(len(I)))
    #st.write(range(len(T)))
    
    a = {}    
    for i in range(len(I)):
        for t in range(len(T)):
            a[i, t] = data_nyuuryoku.iloc[i, t]
    #st.write(a[0, 7])
    
    old_x = {}
    for j in range(len(J)):
        for t in range(len(T)):
            old_x[j, t] = data_shift.iloc[j, t]
    #st.write(old_x[3, 7])
            
#空問題の作成
    model = LpProblem('Shift', sense=LpMinimize)

#決定変数の作成
    x = {}
    for i in range(len(I)):
        for j in range(len(J)):
            for t in range(len(T)):
                x[i, j, t] = LpVariable(f'x{i},{j},{t}', cat=LpBinary)
    v = {}
    for i in range(len(I)):
        v[i] = LpVariable(f'v{i}', lowBound=0)
    w = {}
    for i in range(len(I)):
        for t in range(len(T)):
            w[i, t] = LpVariable(f'w{i}, {t}', cat=LpBinary)
    z = {}
    for i in range(len(I)):
        for t in range(len(T)):
            z[i, t] = LpVariable(f'z{i}, {t}', cat=LpBinary)
            
#制約条件の追加
#条件①
    for j in range(len(J)):
        for t in range(len(T)):
            model += lpSum(x[i, j, t] for i in range(len(I))) == 1
#条件②
    for i in range(len(I)):
        for t in range(len(T)):
            model += lpSum(x[i, j, t] for j in range(len(J))) <= 1
#条件③
    for i in range(len(I)):
        for t in range(len(T)):
            if a[i, t] == 0:
                model += lpSum(x[i, j, t] for j in range(len(J))) == 0
#条件④
    for i in range(len(I)):
        for t in range(len(T)):
            if a[i, t] == 3:
                model += lpSum(x[i, j, t] for j in range(len(J))) == 0
#条件⑤
    s = (len(J) * len(T)) / len(I)
    for i in range(len(I)):
        model += lpSum(x[i, j, t] for j in range(len(J)) for t in range(len(T))) >= s - v[i]
        model += lpSum(x[i, j, t] for j in range(len(J)) for t in range(len(T))) <= s + v[i]
#条件⑥
    for i in range(len(I)):
        for t in range(len(T)-2):
            model += lpSum(x[i, j, t+1] for j in range(len(J))) + lpSum(x[i, j, t+2] for j in range(len(J))) <= 1 + z[i, t+1]
#条件⑦
    for i in range(len(I)):
        for t in range(len(T)):
            if a[i, t] == 2:
                model += lpSum(x[i, j, t] for j in range(len(J))) == 0 + w[i, t]
#条件⑧
    for j in range(len(J)):
        for t in range(len(T)):
            if old_x[j, t] in member_list:
                model += lpSum(x[member_list.index(old_x[j, t]), j, t]) == 1

#目的関数の設定
    model += lpSum(v[i] for i in range(len(I))) + lpSum(z[i, t] for i in range(len(I)) for t in range(len(T))) + lpSum(w[i, t] for i in range(len(I)) for t in range(len(T)))

#最適化の実行
    model.solve()
#    new_data_shift = pd.DataFrame(index=range(4), columns=range(len(T)))
 #   new_data_shift.fillna(0, inplace=True)
    
    if st.button('欠席者対応シフト作成'): 
        status = st.empty()
        bar = st.progress(0)
        for i in range(100):
            status.text(f'Status {i+1}')
            bar.progress(i + 1)
            time.sleep(0.05)
            
        if LpStatus[model.status] == 'Optimal':
            for i in range(len(I)):
                for j in range(len(J)):
                    for t in range(len(T)):
                        if value(x[i, j, t]) > 0.01:
                           # if old_x[j, t] in member_list:
                            #    new_data_shift.iloc[j, t] = old_x[j, t]
                            #else:
                            data_shift.iloc[j, t] = member_list[i]
            st.dataframe(data_shift)
        else:
            st.write('シフトが作成できませんでした')