import streamlit as st
from pulp import *
from openpyxl import *
import pandas as pd
import time


#Excelシートの読み込み
st.session_state.book = load_workbook('卒業研究_シフトデータ.xlsx')
sh1 = st.session_state.book['入力データ']
sh2 = st.session_state.book['シフト表']

#初期情報の入力
st.title('ライブ裏方シフトのシフト表')
st.header('▼ 初期情報の入力')

member = st.number_input('人数を決めてください', min_value=0, max_value=100)
st.write('→ 人数は', member, '人です。')
slot = st.number_input('出演バンド数を決めてください', min_value=0, max_value=100)
st.write('→ 出演バンドは', slot, 'バンドです。')

#定数用のデータの作成
st.session_state.I = [i+1 for i in range(member)]
st.session_state.J = [i+1 for i in range(4)]
st.session_state.T = [i+1 for i in range(slot)]

@st.cache_resource
def input_list_member():
    st.session_state.lst_member = []
    return st.session_state.lst_member

@st.cache_resource
def input_list_turn():
    lst_turn = []
    return lst_turn

#メンバーの名前入力
if int(member) > 0:
    st.session_state.lst_member = input_list_member()
    lst_turn = input_list_turn()
    with st.form('input_name', clear_on_submit=True):
        st.write('メンバーの名前を入力してください')
        member_name = st.text_input('↓↓↓')
        submitted = st.form_submit_button('メンバー追加')
        if submitted:
            st.session_state.lst_member.append(member_name)
            lst_turn.append(0)
        st.table(st.session_state.lst_member)
    if st.button('clear member'):
        input_list_member.clear()
        st.session_state.lst_member = input_list_member()
        
    if len(lst_turn) > 0:
        with st.form('input_turn', clear_on_submit=False):
            if lst_turn.count(0) > 0:
                st.write(st.session_state.lst_member[lst_turn.index(0)], 'が何バンド目に出演するか入力してください')
            elif lst_turn.count(0) == 0:
                st.write('入力完了')
            member_turn = st.number_input('↓↓↓', min_value=1, max_value=slot)
            submitted = st.form_submit_button('出演情報追加')
            if submitted:
                lst_turn[lst_turn.index(0)] = member_turn
            st.table(lst_turn)
        if st.button('clear turn'):
            input_list_turn.clear()
            lst_turn = input_list_turn()
    
    
#Excel書き込み
    if len(lst_turn) > 0:
        st.header('▼ 入力データ保存')
        if st.button('データ保存'):
            status = st.empty()
            bar = st.progress(0)
            for i in range(100):
                status.text(f'Status {i+1}')
                bar.progress(i + 1)
                time.sleep(0.05)
            '→ データの保存が完了しました！'

            for i in range(0, int(slot)):
                sh1.cell(row=1, column=2+i).value = i+1
                sh2.cell(row=1, column=2+i).value = i+1
            for i in range(0, len(st.session_state.lst_member)):
                sh1.cell(row=2+i, column=1).value = st.session_state.lst_member[i]
                for t in range(0, int(slot)):
                    sh1.cell(row=2+i, column=2+t).value = 1
            for i in range(0, len(lst_turn)):
                sh1.cell(row=2+i, column=1+lst_turn[i]).value = 0
            for i in range(0, len(lst_turn)):
                if lst_turn[i] < 8:
                    sh1.cell(row=2+i, column=2+lst_turn[i]).value = 2
            for i in range(0, len(lst_turn)):
                if lst_turn[i] > 1:
                    sh1.cell(row=2+i, column=lst_turn[i]).value = 3
            st.session_state.book.save('卒業研究_シフトデータ_1.xlsx')
            df = pd.read_excel('卒業研究_シフトデータ_1.xlsx', sheet_name='入力データ', index_col=0)
            st.dataframe(df)