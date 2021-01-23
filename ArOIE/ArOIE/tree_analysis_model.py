import xlrd
import numpy as np
import pronoun_link as pl
import sentence_types as st
from xlutils.copy import copy

Head = None
Root_Id = None
UPosTag = None
DepRel = None
MaxPartIndex = None
A2 = None
C = None
SenType = None


def tree_analysis(ArrSen, FileName, SheetN, output_excel_file_name, output_sheet_name):
    Sheet = xlrd.open_workbook(FileName).sheet_by_name(SheetN)
    PartArr = []
    OldPartArr = []
    S_ID_Old = None
    ArrSen = ArrSen[:-1]
    V_Rel = []
    V_Index = -1
    All = None
    # reverse the array
    reverse_array = ArrSen[::-1]
    SenN = int(ArrSen[0]) - 2
    # Print the sentence id and text to the excel file 
    if str(Sheet.cell((int(ArrSen[0]) - 1), 0).value) != "1-2":
        sent_id = str(Sheet.cell(SenN, 0))
        sent_text = str(Sheet.cell(SenN + 1, 0))
        write_to_excel_file(sent_id,sent_text,output_excel_file_name,output_sheet_name)
    else:
        sent_id = str(Sheet.cell((int(ArrSen[0]) - 3), 0))
        sent_text = str(Sheet.cell((int(ArrSen[0]) - 2), 0))
        write_to_excel_file(sent_id,sent_text,output_excel_file_name,output_sheet_name)
    Check = 0
    Verbs = []
    # Extract POS verbs from sentence
    for r in reverse_array :
        UPosTag = Sheet.cell(r, 3).value
        if UPosTag == "VERB" and r not in Verbs :
            Verbs.append(r)
    for i in reverse_array:
        V = None
        if i not in OldPartArr:
            for m in reverse_array:
                UPosTag = Sheet.cell(m, 3).value
                S = None
                O = None
                A1 = None
                V_ID = None
                All_Arr = []
                if UPosTag == "VERB":
                    V_Index = V_Index + 1
                    PartArr.append(m)
                    MaxPartIndex = m
                    CV = 0
                    CS = 0
                    CO = 0
                    S_ID_Pre = None
                    O_ID_Pre = None
                    I = None
                    E = None
                    S_ID = None
                    O_ID = None
                    A1_ID = None
                    S_Arr = []
                    O_Arr = []
                    A1_Arr = []
                    SS = 200
                    out = []
                    for t in reverse_array:
                        if V_Index + 1 < len(Verbs) and t > MaxPartIndex and t not in OldPartArr:
                            if int(Sheet.cell(t, 6).value) == int(Sheet.cell(Verbs[V_Index + 1], 0).value):
                                out.append(t)
                        if t > MaxPartIndex and t not in out:
                            PartArr.append(t)
                            PartArr = sorted(PartArr)
                    for a in PartArr:
                        UPosTag = Sheet.cell(a, 3).value
                        UPosTagPre = Sheet.cell(a - 1, 3).value
                        DepRel = Sheet.cell(a, 7).value
                        XPosTag = str(Sheet.cell(a, 4).value)
                        if UPosTag == "VERB" and CV == 0:
                            V = str(Sheet.cell(a, 1).value)
                            V_ID = a
                            CV = 1
                            V_XPosTag = str(Sheet.cell(a, 4).value)
                        for k in ArrSen:
                            if str(Sheet.cell(V_ID, 0).value) == str(Sheet.cell(k, 6).value) and k not in V_Rel and S == None:
                                V_Rel.append(k)
                            if str(Sheet.cell(V_ID, 6).value) == str(Sheet.cell(k, 0).value) and S == None:
                                SS = k
                        if a != S_ID_Old and CS == 0:
                            if DepRel == "nsubj":
                                S = str(Sheet.cell(a, 1).value)
                                SPre = str(Sheet.cell(a - 1, 1).value)
                                S = pl.pronoun_link(UPosTag, S, SPre, UPosTagPre)
                                S_ID = a
                                S_ID_Pre = int(a - 1)
                                S_ID_Old = a
                                CS = 1
                                S_Head = Sheet.cell(a, 6).value
                                S_Arr_ID = [Sheet.cell(S_ID, 0).value]
                                S_All = S
                        if DepRel == "obj" or DepRel == "obl:arg" and CO == 0 and a != S_ID_Pre:
                            O = str(Sheet.cell(a, 1).value)
                            OPre = str(Sheet.cell(a - 1, 1).value)
                            O = pl.pronoun_link(UPosTag, O, OPre, UPosTagPre)
                            O_ID = a
                            O_ID_Pre = int(a - 1)
                            CO = 1
                    for d in PartArr:
                        UPosTagd = Sheet.cell(d, 3).value
                        UPosTagPred = Sheet.cell(d - 1, 3).value
                        DepReld = Sheet.cell(d, 7).value
                        if DepReld == "case" or DepReld == "obl" or DepReld == "nsubj" :
                            if A1 == None and d != S_ID and d != O_ID_Pre and str(Sheet.cell(d, 3).value) != "CCONJ":
                                A1 = str(Sheet.cell(d, 1).value)
                                APre = str(Sheet.cell(d - 1, 1).value)
                                A1 = pl.pronoun_link(UPosTagd, A1, APre, UPosTagPred)
                                A1_ID = d
                    for h in V_Rel:
                        DepRel = Sheet.cell(h, 7).value
                        if DepRel == "nsubj" and S == None:
                            S = str(Sheet.cell(h, 1).value)
                            S_ID = h
                            S_ID_Pre = int(h - 1)
                            SPre = str(Sheet.cell(h - 1, 1).value)
                            S = pl.pronoun_link(UPosTag, S, SPre, UPosTagPre)
                            V_Rel = []
                            break
                        if str(Sheet.cell(SS, 3).value) == "NOUN" and S == None and SS != 200:
                            S = str(Sheet.cell(SS, 1).value)
                            V_Rel = []
                            break
                    S_Arr = [S_ID]
                    O_Arr = [O_ID]
                    A1_Arr = [A1_ID]
                    for l in PartArr:
                        if S_ID is not None:
                            if Sheet.cell(l, 6).value == Sheet.cell(S_ID, 0).value and l != V_ID and l != O_ID and l != A1_ID and l not in S_Arr:
                                S_Arr.append(l)
                                S_Arr = sorted(S_Arr)
                                S_All = ""
                                for e in S_Arr:
                                    if str(Sheet.cell(e - 1, 1).value) == "ال":
                                        S_All = S_All + " " + "ال" + str(Sheet.cell(e, 1).value)
                                    else:
                                        S_All = S_All + " " + str(Sheet.cell(e, 1).value)
                        if O_ID is not None:
                            if Sheet.cell(l, 6).value == Sheet.cell(O_ID, 0).value and l != V_ID and l != S_ID and l != A1_ID :
                                O_Arr.append(l)
                                O_Arr = sorted(O_Arr)
                                O_All = ""
                                for e in O_Arr:
                                    if str(Sheet.cell(e - 1, 1).value) == "ال":
                                        O_All = O_All + " " + "ال" + str(Sheet.cell(e, 1).value)
                                    else:
                                        O_All = O_All + " " + str(Sheet.cell(e, 1).value)
                        if A1_ID is not None:
                            if Sheet.cell(l, 6).value == Sheet.cell(A1_ID, 0).value and l != V_ID and l != S_ID and l != O_ID:
                                A1_Arr.append(l)
                                A1_Arr = sorted(A1_Arr)
                                A1_All = ""
                                for e in A1_Arr:
                                    if str(Sheet.cell(e - 1, 1).value) == "ال":
                                        A1_All = A1_All + " " + "ال" + str(Sheet.cell(e, 1).value)
                                    else:
                                        A1_All = A1_All + " " + str(Sheet.cell(e, 1).value)
                    if V_ID is not None:
                        All_Arr.append(V_ID)
                    if S_ID is not None :
                        All_Arr = All_Arr + S_Arr
                    if O_ID is not None :
                        All_Arr = All_Arr + O_Arr
                    if A1_ID is not None :
                        All_Arr = All_Arr + A1_Arr
                    All = ""
                    for e in All_Arr :
                        if str(Sheet.cell(e - 1, 1).value) == "ال" :
                            All = All + " " + "ال" + str(Sheet.cell(e, 1).value)
                        else :
                            All = All + " " + str(Sheet.cell(e, 1).value)
                    SenType = str(st.SenType(V, S, O, A1))
                    print("[+] Sentence Type is" + " " + SenType)
                    V_Rel = []
                    workbook = xlrd.open_workbook(output_excel_file_name)
                    sheet = workbook.sheet_by_name(output_sheet_name)
                    wb = copy(workbook)
                    sheetw = wb.get_sheet(0)
                    n = sheet.nrows
                    sheetw.write(n + 1, 0, st.SenType(V, S, O, A1))
                    wb.save(output_excel_file_name)
                elif m == ArrSen[0] and V is None and Check == 0 :
                    CI = 0
                    CE = 0
                    for a in ArrSen :
                        UPosTag = Sheet.cell(a, 3).value
                        UPosTagPre = Sheet.cell(a - 1, 3).value
                        XPosTag = Sheet.cell(a, 4).value
                        DepRel = Sheet.cell(a, 7).value
                        Check = 1
                        if a == ArrSen[0] and CI == 0 :
                            I = str(Sheet.cell(a, 1).value)
                            CI = 1
                        if DepRel == "root" and CE == 0 :
                            E = str(Sheet.cell(a, 1).value)
                            CE = 1
                    print("[*] Sentence Type is" + " " + str(st.NSenType(I, E)))
                    workbook = xlrd.open_workbook(output_excel_file_name)
                    sheet = workbook.sheet_by_name(output_sheet_name)
                    wb = copy(workbook)
                    sheetw = wb.get_sheet(0)
                    n = sheet.nrows
                    sheetw.write(n + 1, 0, st.NSenType(I, E))
                    wb.save(output_excel_file_name)
                reverse_array = list(set(reverse_array) - set(PartArr))
                reverse_array = reverse_array[: :-1]
                OldPartArr = PartArr
                PartArr = []

def write_to_excel_file(sent_id,sent_text,excel_fname,sheet_name):
  workbook = xlrd.open_workbook(excel_fname)
  sheetk = workbook.sheet_by_name(sheet_name)
  wb = copy(workbook)
  sheetw = wb.get_sheet(0)
  n = sheetk.nrows
  sheetw.write(n + 1, 0, sent_id)
  sheetw.write(n + 2, 0, sent_text)
  wb.save(excel_fname)
