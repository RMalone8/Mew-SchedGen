import customtkinter as ctk
from PIL import Image
from openpyxl import load_workbook
import random
import shutil
from pathlib import Path
import numpy as np
from datetime import datetime
from template_create import template_generator

# Sets the appearance mode of the application
# "System" sets the appearance same as that of the system
ctk.set_appearance_mode("System")        
 
# Sets the color of the widgets 
# Supported themes: green, dark-blue, blue
ctk.set_default_color_theme("dark-blue")
 
appWidth=1275
appHeight=650
errormessage="Click Blue Button to Generate Schedule"
activity1="Soccer @ Complex"
activity2="Movie @ Nook" # Double
activity3="Movie @ Video Theatre" # Double
activity4="Activity Shuffle 1"
activity5="Activity Shuffle 2"
activity6="FF @ Complex"
activity7="Games @ Dance Studio" # Girls
activity8="Games @ Vault"
activity9="Downball Tournament @ Garden" # Boys
activity10="Bracelet Making @ Pavillion" # Girls
activity11="Rainy Day Hike @ PA"
activity12="Tennis @ Tennis Center"
activity_names=[activity1,activity2,activity3,activity4,activity5,activity6,activity7,activity8,activity9,activity10,activity11,activity12]

class MyFrame(ctk.CTkScrollableFrame):
    def __init__(self, master, actnames: list, **kwargs):
        super().__init__(master, **kwargs)
        #Initialize Checkboxes

        self.checkboxVars = [ctk.StringVar(value="off") for i in range(len(actnames))]

        self.checkboxSpan = [ctk.StringVar(value="off") for i in range(len(actnames))]

        self.checkboxboy = [ctk.StringVar(value="on") for i in range(len(actnames))]

        self.checkboxgirl = [ctk.StringVar(value="on") for i in range(len(actnames))]

        self.checkboxsimul= [ctk.StringVar(value="off") for i in range(len(actnames))]

        self.choices = [ctk.CTkCheckBox(self, text=actnames[i], variable=self.checkboxVars[i], onvalue="on", offvalue="off") for i in range(len(actnames))]
        for i in range(len(self.choices)):
            self.choices[i].grid(row=i, column=0, padx=20, pady=20, sticky="ew")

        self.activity_entries = []
        self.activity_buttons = []
        
        for i, actname in enumerate(actnames):
            self.change_activity_widgets(actname, i)

        for i in range(len(self.checkboxSpan)):
            self.checkboxSpan[i] = ctk.CTkCheckBox(self, text="Span Two Periods", variable=self.checkboxSpan[i], onvalue="on", offvalue="off")
            self.checkboxSpan[i].grid(row=i, column=3, padx=20, pady=20, sticky="ew")

        for i in range(len(self.checkboxsimul)):
            self.checkboxsimul[i] = ctk.CTkCheckBox(self, text="Simultaneous", variable=self.checkboxsimul[i], onvalue="on", offvalue="off")
            self.checkboxsimul[i].grid(row=i, column=4, padx=20, pady=20, sticky="ew")

        for i in range(len(self.checkboxgirl)):
            self.checkboxgirl[i] = ctk.CTkCheckBox(self, text="Girl's Side", variable=self.checkboxgirl[i], onvalue="on", offvalue="off")
            self.checkboxgirl[i].grid(row=i, column=5, padx=20, pady=20, sticky="ew")

        for i in range(len(self.checkboxboy)):
            self.checkboxboy[i] = ctk.CTkCheckBox(self, text="Boy's Side", variable=self.checkboxboy[i], onvalue="on", offvalue="off")
            self.checkboxboy[i].grid(row=i, column=6, padx=20, pady=20, sticky="ew")

    def change_activity_widgets(self, placeholder_text, row):
        entry = ctk.CTkEntry(self, placeholder_text=placeholder_text)
        entry.grid(row=row, column=1, columnspan=1, padx=10, pady=10, sticky="ew")

        def change_activity():
            self.choices[row].configure(text=entry.get())
        
        button = ctk.CTkButton(self, text="Change", text_color="#000000", corner_radius=32, fg_color="#FFFFFF", hover_color="#0047AB", command=change_activity)
        button.grid(row=row, column=2, columnspan=1, padx=10, pady=10, sticky="ew")
        
        self.activity_entries.append(entry)
        self.activity_buttons.append(button)

# Create App class
class App(ctk.CTk):
# Layout of the GUI will be written in the init itself
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.title("Mew: Schedule Generator")    

        self.geometry(f"{appWidth}x{appHeight}")    
 
        num_of_acts = 12 # Initial number of actvities!
        self.my_frame = MyFrame(master=self, actnames=activity_names, width=1200, height=300)
        self.my_frame.grid(row=2, column=0,columnspan=12, padx=20, pady=20)


        # Schedule Name/Title
        self.nameLabel = ctk.CTkLabel(self,
                                      text="Schedule Name")
        self.nameLabel.grid(row=0, column=0,
                            padx=3, pady=10,
                            sticky="ew")
        self.nameEntry = ctk.CTkEntry(self,
                         placeholder_text="Enter Schedule Name")
        self.nameEntry.grid(row=0, column=1,
                            columnspan=3, padx=3,
                            pady=10, sticky="ew")



        #Alter Number of Activities
        def alternumacts():
            num_of_acts=int(self.numberactsentry.get())
            if(num_of_acts>=12 and num_of_acts<20): # 20 is arbitrary right now
                # Reinitialize the scroll!!
                actnames = [self.my_frame.choices[i].cget("text") if i < len(self.my_frame.choices) else "FILLER" for i in range(num_of_acts)]
            elif(num_of_acts<12):
                # Tell the user that this number cannot be lower than 12!
                num_of_acts = 12
                print("Yo, not cool dude")

        self.numberacts = ctk.CTkLabel(self,
                                      text="Number of Activities")
        self.numberacts.grid(row=0, column=5,
                            padx=5, pady=10,
                            sticky="ew")
        self.numberactsentry = ctk.CTkEntry(self,
                         placeholder_text="Need at least 12")
        self.numberactsentry.grid(row=0, column=6,
                            columnspan=3, padx=5,
                            pady=10, sticky="ew")
        self.numberactsbutton=ctk.CTkButton(self,text="Change", corner_radius=32, fg_color="#0000FF",hover_color="#33BFFF",command=alternumacts)
        self.numberactsbutton.grid(row=0,column=9,columnspan=1,padx=15,pady=10,sticky="ew")        
         
        def testrun():
            self.nameLabel = ctk.CTkLabel(self,text="An Error Has Occured")
            self.nameLabel.grid(row=14, column=0,padx=10, pady=10,sticky="ew")
        
        def selectall():
            for i in range(len(self.my_frame.choices)):
                self.my_frame.checkboxVars[i].set("on")
            
        def run(): #button press
            opts = [True if self.my_frame.checkboxVars[i].get() == "on" else False for i in range(num_of_acts)]
            actnames=[self.my_frame.choices[i].cget("text") for i in range(len(self.my_frame.choices))]
            act_dict = []
            #Sort out the categories
            categs = ["All" if self.my_frame.checkboxboy[i].get() == self.my_frame.checkboxgirl[i].get() else "JustBoy" if self.my_frame.checkboxboy[i].get() == "on" else "JustGirl" for i in range(num_of_acts)]
            
            categs = [categs[i]+"Simul" if self.my_frame.checkboxsimul[i].get() == "on" else categs[i] for i in range(num_of_acts)]

            categs = [categs[i]+"Double" if self.my_frame.checkboxSpan[i].get() == "on" else categs[i] for i in range(num_of_acts)]

            time_period = "morning" if self.checkboxmorning.get() == "on" else "afternoon" if self.checkboxafternoon.get() == "on" else "wholeday"

            for i in range(len(opts)):
                    if opts[i]:
                        act_dict.append({"name": actnames[i], "type": categs[i]})
            global specs

            if time_period == "morning":
                specs = {"rows": ["4", "5", "6", "7", "8", "10", "11", "12", "13", "14"],
                        "cols": ['C', 'D', 'E'],
                        "blocked_coords": []} 
            elif time_period == "afternoon":
                specs = {"rows": ["4", "5", "6", "7", "8", "10", "11", "12", "13", "14"],
                        "cols": ['D', 'F', 'G', 'H'],
                        "blocked_coords": [[2,0], [2,1], [2,2], [2,5], [2,6], [2,7]]} 
            else:
                specs = {"rows": ["4", "5", "6", "7", "8", "10", "11", "12", "13", "14"],
                        "cols": ['C', 'D', 'E', 'G', 'I', 'J', 'K'],
                        "blocked_coords": [[5,0], [5,1], [5,2], [5,5], [5,6], [5,7]]} 

            def getSchedule(b_groups: int, g_groups: int, activities: list, sheet_specs: dict):
                count = 0
                result = generateSchedule(b_groups= b_groups, g_groups= g_groups, activities= activities, sheet_specs=sheet_specs)
                while not result and count < 100:
                    result = generateSchedule(b_groups= b_groups, g_groups= g_groups, activities= activities, sheet_specs=sheet_specs)
                    count +=1
                return result

            def generateSchedule(b_groups: int, g_groups: int, activities: list, sheet_specs: dict):
                periods = len(sheet_specs["cols"])
                groups = b_groups + g_groups
                new_activity = ""
                global override_coords
                override_coords = []

                book = template_generator(time_period) #load_workbook("schedule_temps/template_sheet_" + time_period + ".xlsx")
                sheet = book.active

                # Checking if a valid schedule can even be made!
                if groups != len(sheet_specs["rows"]):
                    return "Invalid number of rows for number of chosen groups."
                if groups > len(activities) or periods > len(activities):
                    return "Not enough groups or periods to fulfill the activities selected."
                
                global schedule
                schedule = [["" for group in range(groups)] for period in range(periods)]

                for i in range(groups):
                    for j in range(periods):
                        while True:
                            if not schedule[j][i]: # If there isn't already something there
                                new_activity = chooseValidActivity(group=i, period=j, groups=groups, periods=periods, activities=activities)
                                if new_activity:
                                    schedule[j][i] = new_activity["name"]
                                    sheet[sheet_specs["cols"][j]+sheet_specs["rows"][i]].value = new_activity["name"]
                                    break
                                else:
                                    return False
                            else:
                                for coords in override_coords:
                                    if [j, i] == coords["coords"]: # Are these coordinates to be skipped?
                                        sheet[sheet_specs["cols"][j]+sheet_specs["rows"][i]].value = schedule[j][i]
                                        override_coords.remove(coords)
                                        break 
                                break

                # Sending the finished schedule...
                sheet["A1"].value = self.nameEntry.get() 
                book.save('new_sheet.xlsx')
                output_path=Path.home()/'Downloads'/f'Generated_Schedule_{datetime.now().month}_{datetime.now().day}_{datetime.now().year}.xlsx'
                shutil.copyfile('new_sheet.xlsx',output_path) 
                return schedule

            def chooseValidActivity(group: int, period: int, groups: int, periods: int, activities: list) -> str:
                avail_activities = activities.copy()
                for act in avail_activities[:]:
                    if act["name"] in schedule[period]:
                        avail_activities.remove(act)
                    elif act["name"] in np.transpose(schedule)[group]:
                        avail_activities.remove(act)
                    elif "Girl" in act["type"] and group < groups/2:
                        avail_activities.remove(act)
                    elif "Boy" in act["type"] and group >= groups/2:
                        avail_activities.remove(act)

                    #Simulateous Conditions    
                    elif "Simul" in act["type"] and (group == groups - 1 or group == groups/2 - 1):
                        avail_activities.remove(act)
                    elif "Simul" in act["type"] and [period, group+1] in specs['blocked_coords']:
                        avail_activities.remove(act)
                    elif "Simul" in act["type"] and schedule[period][group+1]:
                        avail_activities.remove(act)

                    #Double Conditions
                    elif "Double" in act["type"] and (period == periods - 1 or act["name"] in schedule[period+1]):
                        avail_activities.remove(act)
                    elif "Double" in act["type"] and abs(ord(specs['cols'][period]) - ord(specs['cols'][period+1])) > 1:
                        avail_activities.remove(act)
                    elif "Double" in act["type"] and [period+1, group] in specs['blocked_coords']:
                        avail_activities.remove(act)
                    elif "Double" in act["type"] and schedule[period+1][group]:
                        avail_activities.remove(act)

                if [period, group] in specs['blocked_coords']:
                    new_act = {"name": "", "type": "None"}
                elif avail_activities:
                    new_act = random.choice(avail_activities)
                else:
                    return False
                if "Simul" in new_act["type"] and "Double" in new_act["type"]:
                    schedule[period][group+1] = new_act["name"]
                    override_coords.append({"coords": [period, group+1]})
                    schedule[period+1][group] = new_act["name"]
                    override_coords.append({"coords": [period+1, group]})
                    schedule[period+1][group+1] = new_act["name"]
                    override_coords.append({"coords": [period+1, group+1]})
                elif "Simul" in new_act["type"]:
                    schedule[period][group+1] = new_act["name"]
                    override_coords.append({"coords": [period, group+1]})
                elif "Double" in new_act["type"]:
                    schedule[period+1][group] = new_act["name"]
                    override_coords.append({"coords": [period+1, group]})
                return new_act
            
            schedy = getSchedule(b_groups=5, g_groups=5, activities=act_dict, sheet_specs=specs)
            print(schedy)
            


        # Functions to maintain one enabled timeframe checkbox
        def reset_checkboxes_morn():
            if self.checkboxmorning.get() == "on":
                self.checkboxafternoon.set("off")
                self.checkboxwholeday.set("off")
        def reset_checkboxes_afternoon():
            if self.checkboxafternoon.get() == "on":
                self.checkboxmorning.set("off")
                self.checkboxwholeday.set("off")
        def reset_checkboxes_wholeday():
            if self.checkboxwholeday.get() == "on":
                self.checkboxafternoon.set("off")
                self.checkboxmorning.set("off")


        # Creating the checkboxes
        self.checkboxmorning=ctk.StringVar(value="off")
        self.checkboxafternoon=ctk.StringVar(value="off")
        self.checkboxwholeday=ctk.StringVar(value="on")
        self.choicemorning = ctk.CTkCheckBox(self,
                            text="Morning",
                            variable=self.checkboxmorning,
                            onvalue="on",
                            offvalue="off",
                            command=reset_checkboxes_morn)                               
        self.choicemorning.grid(row=7, column=3,
                          padx=5, pady=5,
                          sticky="ew")
        
        self.choiceafternoon = ctk.CTkCheckBox(self,
                            text="Afternoon",
                            variable=self.checkboxafternoon,
                            onvalue="on",
                            offvalue="off",
                            command=reset_checkboxes_afternoon)                             
        self.choiceafternoon.grid(row=7, column=4,
                          padx=5, pady=5,
                          sticky="ew")

        self.choicewholeday = ctk.CTkCheckBox(self,
                            text="Whole Day",
                            variable=self.checkboxwholeday,
                            onvalue="on",
                            offvalue="off",
                            command=reset_checkboxes_wholeday)                               
        self.choicewholeday.grid(row=7, column=5,
                          padx=5, pady=5,
                          sticky="ew")


        #Make Spreadsheet Button
        self.testButton=ctk.CTkButton(self,text="Select All Activities", text_color="#000000", corner_radius=32, fg_color="#FFFFFF",hover_color="#FFFFFF",command=selectall)
        self.testButton.grid(row=1,column=3,columnspan=3,padx=10,pady=10,sticky="ew")

        self.testButton=ctk.CTkButton(self,text="Select Duration of Schedule", text_color="#000000", corner_radius=32, fg_color="#FFFFFF",hover_color="#FFFFFF",command=testrun)
        self.testButton.grid(row=6,column=3,columnspan=3,padx=10,pady=10,sticky="ew")

        self.generateResultsButton=ctk.CTkButton(self,text="Generate Schedule", corner_radius=32, fg_color="#0000FF",hover_color="#33BFFF",command=run)
        self.generateResultsButton.grid(row=8,column=3,columnspan=3,padx=10,pady=10,sticky="ew")

        self.testButton=ctk.CTkButton(self,text=errormessage, text_color="#000000", corner_radius=32, fg_color="#FFFFFF",hover_color="#FFFFFF",command=testrun)
        self.testButton.grid(row=9,column=3,columnspan=3,padx=10,pady=10,sticky="ew")
      



if __name__ == "__main__":
    app = App()
    # Runs the app
    app.mainloop()   
