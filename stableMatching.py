import openpyxl as xl
import tkinter as tk
import texttable as tt
from PIL import Image, ImageTk
from tkinter.filedialog import askopenfile

root = tk.Tk()
root.resizable(False, False)
root.geometry("800x550")
root.title("SOPP Student Org Matching")

# canvas = tk.Canvas(root, width=800, height=800)
# canvas.grid()


#logo
logo = Image.open('SOPP banners_800x125.png').resize((800, 125))
logo = ImageTk.PhotoImage(logo)
logo_label = tk.Label(image=logo)
logo_label.image = logo
logo_label.grid(column=1, row=0)

#instructions
instructions_text = tk.StringVar()
instructions = tk.Label(root, textvariable=instructions_text)
instructions.grid(column=1, row=1)
instructions_text.set("Please select a .xlsx file as the Student Org preference response sheet")

loc = None
wb = None
ws1 = None
ws2 = None
orgColLetter = None
fairColLetter = None
prefStart1 = None
prefEnd1 = None
prefStart2 = None
prefEnd2 = None
orgCol = None
fairCol = None
orgs = []
fairs = []
orgPref = []
fairPref = []

def open_file():
    browse_text.set("Loading...")
    file = askopenfile(parent=root, mode='rb', title="Choose a file")
    
    if file:
        global loc
        global wb
        loc = file.name
        wb = xl.load_workbook(filename = loc)

        browse_btn.grid_remove()
        input_button.grid(column=1, row=2, pady=5)
        instructions_text.set("Please fill out the information below")
        prompt1.grid(column=1, row=5)
        entry1.grid(column=1, row=6)
        prompt2.grid(column=1, row=7)
        entry2.grid(column=1, row=8, padx=20)
        prompt3.grid(column=1, row=9)
        entry3.grid(column=1, row=10, padx=20)
        prompt4.grid(column=1, row=11)
        entry4.grid(column=1, row=12, padx=20)

        prompt5.grid(column=1, row=13)
        entry5.grid(column=1, row=14)
        prompt6.grid(column=1, row=15)
        entry6.grid(column=1, row=16, padx=20)
        prompt7.grid(column=1, row=17)
        entry7.grid(column=1, row=18, padx=20)
        prompt8.grid(column=1, row=19)
        entry8.grid(column=1, row=20, padx=20)

def setVariables():
    # user input
    global orgColLetter
    global orgCol 
    global prefStart1
    global prefEnd1
    global ws1
    global ws2
    global fairCol
    global fairColLetter
    global prefStart2
    global prefEnd2

    ws1 = wb[entry1.get()]
    ws2 = wb[entry5.get()]
    orgColLetter = entry2.get()
    prefStart1 = entry3.get()
    prefEnd1 = entry4.get()
    fairColLetter = entry6.get()
    prefStart2 = entry7.get()
    prefEnd2 = entry8.get()
    orgCol = ws1[orgColLetter]
    fairCol = ws2[fairColLetter]

    for org in orgCol:
        value = org.value
        if (not(value is None)):
            orgs.append(value)
    orgs.pop(0)

    for fair in fairCol:
        value = fair.value
        if (not(value is None)):
            fairs.append(value)
    fairs.pop(0)

    for i in range(2, len(orgs) + 2):
        indiv = [ws1[orgColLetter + str(i)].value]
        for j in ws1[prefStart1 + str(i) : prefEnd1 + str(i)][0]:
            indiv.append(j.value)
        orgPref.append(indiv)

    for i in range(2, len(fairs) + 2):
        indiv = [ws2[fairColLetter + str(i)].value]
        for j in ws2[prefStart2 + str(i) : prefEnd2 + str(i)][0]:
            indiv.append(j.value)
        fairPref.append(indiv)
    
    def fPrefersS1OverS(fairPref, f, s, s1):
        results = [True, True, True]
        for tentativeOrg in s1:
            for i in range(len(fairPref[0])):
                if (fairPref[fairs.index(f)][i] == s):
                    results[s1.index(tentativeOrg)] = False
                elif (fairPref[fairs.index(f)][i] == tentativeOrg):
                    break

        check_if_org_in_list = [org in fairPref[fairs.index(f)] for org in s1]
        if (False in check_if_org_in_list and False in results):
                results = check_if_org_in_list
        return results

    # orgs 'propose' to fairs -> org optimal
    def stableMatching(prefOrg):
        # stores the student org partners for each fair
        fPartner = [["None", "None", "None"] for i in range(len(fairs))]

        oFree = [True for i in range(len(orgs))]
        oCurrChoice = [1 for i in range(len(orgs))]

        numOfFreeOrgs = len([1 for i in oFree if i == True])

        # represents the iteration day the algorithm is on
        iter = 1

        # algorithm terminates after every org has had the chance to propose to every fair on their preference list
        while (iter < (len(orgs) * len(orgs) - 2 * len(orgs)) and numOfFreeOrgs > 0):
            for org in orgs:
                currOrgIndex = orgs.index(org)
                # if it's the case that the current student org is already matched,
                # then continue to the next org
                if (oFree[currOrgIndex] == False):
                    continue
                #if not, then propose to their next best choice
                topChoice = prefOrg[currOrgIndex][oCurrChoice[currOrgIndex]] 
                # if no one has proposed to the student org's top choice, then propose
                if ("None" in fPartner[fairs.index(topChoice)]):
                    oFree[currOrgIndex] = False
                    fPartner[fairs.index(topChoice)][fPartner[fairs.index(topChoice)].index("None")] = org
                # 3 orgs have already proposed and we need to compare preferences
                else:
                    tentativeOrgs = fPartner[fairs.index(topChoice)]
                    # if the fair prefers the current org to one of the tentative orgs
                    newPreferences = fPrefersS1OverS(fairPref, topChoice, s=orgs[currOrgIndex], s1=tentativeOrgs) 
                    if (False in newPreferences):
                        replacedOrg = tentativeOrgs[newPreferences.index(False)]
                        oFree[orgs.index(replacedOrg)] = True
                        oCurrChoice[orgs.index(replacedOrg)] += 1 
                        # check to see if we're already gone over all of the org's preferences
                        if (oCurrChoice[orgs.index(replacedOrg)] == 4):
                            oFree[orgs.index(replacedOrg)] = False
                            oCurrChoice[orgs.index(replacedOrg)] -= 1

                        fPartner[fairs.index(topChoice)][fPartner[fairs.index(topChoice)].index(replacedOrg)] = org
                        oFree[currOrgIndex] = False
                    else:
                        oCurrChoice[currOrgIndex] += 1 
                        if (oCurrChoice[currOrgIndex] == 4):
                            oFree[currOrgIndex] = False
                            oCurrChoice[currOrgIndex] -= 1
                numOfFreeOrgs = len([1 for i in oFree if i == True])
            iter += 1
        
        result = []
        for i in range(len(fairs)):
            fPartner[i].insert(0, fairs[i])
            result.append(fPartner[i])
        return result

    stuPref = list(orgPref)
    outcome = stableMatching(stuPref)
    outcome.insert(0, ["Event", "Choice 1", "Choice 2", "Choice 3"])
    table = tt.Texttable()
    table.set_cols_align(["l", "l", "l", "l"])
    table.set_cols_valign(["m", "m", "m", "m"])
    table.add_rows(outcome)
    result_box.insert(1.0, table.draw())

    instructions_text.set("Here are the results.")
    input_button.grid_remove()
    prompt1.grid_remove()
    prompt2.grid_remove()
    prompt3.grid_remove()
    prompt4.grid_remove()
    prompt5.grid_remove()
    prompt6.grid_remove()
    prompt7.grid_remove()
    prompt8.grid_remove()
    entry1.grid_remove()
    entry2.grid_remove()
    entry3.grid_remove()
    entry4.grid_remove()
    entry5.grid_remove()
    entry6.grid_remove()
    entry7.grid_remove()
    entry8.grid_remove()
    result_box.grid(column=1, row=2)

#browse button 
browse_text = tk.StringVar()
browse_btn = tk.Button(root, textvariable=browse_text, command=lambda:open_file(), bg='#505c8c', fg="white", height=2, width=15)
browse_text.set("Browse")
browse_btn.grid(column=1, row=2)

#Enter input button
input_text = tk.StringVar()
input_button = tk.Button(root, textvariable=input_text, command=lambda:setVariables(), bg='#505c8c', fg="white", height=2, width=15)
input_text.set("Confirm")

#prompt buttons
prompt1 = tk.Label(root, text="What's the name of the sheet where Student Org preferences are?")
entry1 = tk.Entry(root, width=25)
prompt2 = tk.Label(root, text="The column letter where Student Org names are located:")
entry2 = tk.Entry(root, width=25)
prompt3 = tk.Label(root, text="The column letter where org preferences begin:")
entry3 = tk.Entry(root, width=25)
prompt4 = tk.Label(root, text="The column letter where org preferences end:")
entry4 = tk.Entry(root, width=25)
prompt5 = tk.Label(root, text="What's the name of the sheet where Fair Org preferences are?")
entry5 = tk.Entry(root, width=25)
prompt6 = tk.Label(root, text="The column letter where Fair names are located:")
entry6 = tk.Entry(root, width=25)
prompt7 = tk.Label(root, text="The column letter where Fair preferences begin:")
entry7 = tk.Entry(root, width=25)
prompt8 = tk.Label(root, text="The column letter where Fair preferences end:")
entry8 = tk.Entry(root, width=25)

#results text box
result_box = tk.Text(root, height=20, width=80, pady=15, padx=15)

root.mainloop()