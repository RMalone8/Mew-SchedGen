import customtkinter as ctk
from PIL import Image
import openpyxl as pyxl
from openpyxl import Workbook
import random
import shutil
from pathlib import Path
import numpy as np
from datetime import datetime
from template_create import template_generator
import os

# Sets the appearance mode of the application
# "System" sets the appearance same as that of the system
ctk.set_appearance_mode("System")        
 
# Sets the color of the widgets 
# Supported themes: green, dark-blue, blue
ctk.set_default_color_theme("dark-blue")
 
appWidth=1275
appHeight=650

activity_names = []

if os.path.exists("past_activity_data.xlsx"):
    print("Here")
    data_book = pyxl.open("past_activity_data.xlsx")
    data_sheet = data_book.active
    for i in range(2, 2+data_sheet["A1"].value):
        activity_names.append(data_sheet[f"A{i}"].value)
    data_book.save("past_activity_data.xlsx")
else:
    data_book = Workbook()
    data_sheet = data_book.active

errormessage="Click Blue Button to Generate Schedule"
# activity_names=[activity1,activity2,activity3,activity4,activity5,activity6,activity7,activity8,activity9,activity10,activity11,activity12]
extra_act = ["N/A" for _ in range(30 - len(activity_names))]
activity_names += extra_act

# Will be loaded in from data_sheet as well, but this is for testing
previous_schedules = [[['FF @ Complex', 'Downball Tournament @ Garden', 'Games @ Vault', 'Movie @ Video Theatre', 'Games @ Dance Studio', 'Rainy Day Hike @ PA', 'Bracelet Making @ Pavillion', 'Tennis @ Tennis Center', 'Movie @ Nook', 'Activity Shuffle 2'], ['Games @ Vault', 'Movie @ Video Theatre', 'Tennis @ Tennis Center', 'Downball Tournament @ Garden', 'Soccer @ Complex', 'Games @ Dance Studio', 'Movie @ Nook', 'Activity Shuffle 1', 'FF @ Complex', 'Rainy Day Hike @ PA'], ['Rainy Day Hike @ PA', 'Movie @ Nook', 'Games @ Dance Studio', 'Bracelet Making @ Pavillion', 'Soccer @ Complex', 'FF @ Complex', 'Tennis @ Tennis Center', 'Activity Shuffle 2', 'Movie @ Video Theatre', 'Downball Tournament @ Garden'], ['Activity Shuffle 1', 'Tennis @ Tennis Center', 'Movie @ Nook', 'Games @ Dance Studio', 'Bracelet Making @ Pavillion', 'Activity Shuffle 2', 'Games @ Vault', 'Movie @ Video Theatre', 'Rainy Day Hike @ PA', 'FF @ Complex'], ['Games @ Dance Studio', 'Games @ Vault', 'Activity Shuffle 2', 'Activity Shuffle 1', 'Rainy Day Hike @ PA', 'Movie @ Nook', 'Movie @ Video Theatre', 'Downball Tournament @ Garden', 'Bracelet Making @ Pavillion', 'Tennis @ Tennis Center'], ['', '', '', 'Games @ Vault', 'FF @ Complex', '', '', '', 'Tennis @ Tennis Center', 'Movie @ Video Theatre'], ['Tennis @ Tennis Center', 'Games @ Dance Studio', 'FF @ Complex', 'Movie @ Nook', 'Downball Tournament @ Garden', 'Bracelet Making @ Pavillion', 'Activity Shuffle 2', 'Rainy Day Hike @ PA', 'Games @ Vault', 'Activity Shuffle 1']],
                      [['Downball Tournament @ Garden', 'Activity Shuffle 1', 'Movie @ Video Theatre', 'Basketball @ The Spot', 'Games @ Vault', 'Activity Shuffle 2', 'Games @ Dance Studio', 'Movie @ Nook', 'FF @ Complex', 'Bracelet Making @ Pavillion'], ['Movie @ Video Theatre', 'Movie @ Nook', 'Activity Shuffle 1', 'Basketball @ The Spot', 'Bracelet Making @ Pavillion', 'Games @ Dance Studio', 'FF @ Complex', 'Tennis @ Tennis Center', 'Soccer @ Complex', 'Games @ Vault'], ['Games @ Dance Studio', 'Soccer @ Complex', 'Activity Shuffle 2', 'FF @ Complex', 'Movie @ Nook', 'Bracelet Making @ Pavillion', 'Downball Tournament @ Garden', 'Movie @ Video Theatre', 'Rainy Day Hike @ PA', 'Activity Shuffle 1'], ['Movie @ Nook', 'Activity Shuffle 2', 'Tennis @ Tennis Center', 'Bracelet Making @ Pavillion', 'Activity Shuffle 1', 'Downball Tournament @ Garden', 'Rainy Day Hike @ PA', 'Games @ Dance Studio', 'Movie @ Video Theatre', 'FF @ Complex'], ['Rainy Day Hike @ PA', 'Games @ Vault', 'Bracelet Making @ Pavillion', 'Movie @ Nook', 'Soccer @ Complex', 'FF @ Complex', 'Movie @ Video Theatre', 'Downball Tournament @ Garden', 'Activity Shuffle 2', 'Games @ Dance Studio'], ['', '', '', 'Rainy Day Hike @ PA', 'Downball Tournament @ Garden', '', '', '', 'Bracelet Making @ Pavillion', 'Tennis @ Tennis Center'], ['FF @ Complex', 'Rainy Day Hike @ PA', 'Soccer @ Complex', 'Activity Shuffle 2', 'Tennis @ Tennis Center', 'Movie @ Nook', 'Activity Shuffle 1', 'Bracelet Making @ Pavillion', 'Games @ Dance Studio', 'Movie @ Video Theatre']]]

class MyFrame(ctk.CTkScrollableFrame):
    def __init__(self, master, actnames: list, **kwargs):
        super().__init__(master, **kwargs)
        #Initialize Checkboxes

        self.checkboxVars = [ctk.StringVar(value="off") for _ in range(len(actnames))]

        self.checkboxSpan = [ctk.StringVar(value="off") for _ in range(len(actnames))]

        self.checkboxboy = [ctk.StringVar(value="on") for _ in range(len(actnames))]

        self.checkboxgirl = [ctk.StringVar(value="on") for _ in range(len(actnames))]

        self.checkboxsimul= [ctk.StringVar(value="off") for _ in range(len(actnames))]

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

class Memory(ctk.CTkScrollableFrame):
    def __init__(self, master: list, **kwargs):
        super().__init__(master, **kwargs)

        def save_previous_schedule():
            temp=1
            self.save_schedule_button=ctk.CTkButton(self,text="Save Schedule", corner_radius=32, fg_color="#0000FF",hover_color="#33BFFF",command=save_previous_schedule)
            self.save_schedule_button.grid(row=0,column=0,columnspan=1,padx=10,pady=10,sticky="ew")


# Create App class
class App(ctk.CTk):
# Layout of the GUI will be written in the init itself
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.title("Mew: Schedule Generator")    

        self.geometry(f"{appWidth}x{appHeight}")    
 
        self.num_of_acts = 12 # Initial number of actvities!

        self.my_frame = MyFrame(master=self, actnames=activity_names, width=1200, height=300)
        self.my_frame.grid(row=2, column=0,columnspan=12, padx=20, pady=20)

        #self.memory=Memory(master=self, width=200, height=200)
        #self.memory.grid(row=7, column=8, columnspan=12, padx=20, pady=20)


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
            self.num_of_acts=int(self.numberactsentry.get())
            if(int(self.numberactsentry.get())<12):
                # Tell the user that this number cannot be lower than 12!
                print("Yo, not cool dude")
            elif(int(self.numberactsentry.get())>30):
                print("Too high!")
            else: 
                self.num_of_acts = int(self.numberactsentry.get())
                for i in range(self.num_of_acts):
                    self.my_frame.checkboxVars[i].set("on")

        self.numberacts = ctk.CTkLabel(self,
                                       text="Number of Activities")
        self.numberacts.grid(row=0, column=5,
                             padx=5, pady=10,
                             sticky="ew")
        self.numberactsentry = ctk.CTkEntry(self,
                         placeholder_text="Minimum of 12")
        self.numberactsentry.grid(row=0, column=6,
                            columnspan=1, padx=5,
                            pady=10, sticky="ew")
        self.numberactsbutton=ctk.CTkButton(self,text="Select", corner_radius=32, fg_color="#0000FF",hover_color="#33BFFF",command=alternumacts)
        self.numberactsbutton.grid(row=0,column=7,columnspan=1,padx=15,pady=10,sticky="ew")        
         
        def testrun():
            self.nameLabel = ctk.CTkLabel(self,text="An Error Has Occured")
            self.nameLabel.grid(row=14, column=0,padx=10, pady=10,sticky="ew")
        
        def deselectall():
            for i in range(len(self.my_frame.choices)):
                self.my_frame.checkboxVars[i].set("off")
            
        def run(): #button press

            opts = [True if self.my_frame.checkboxVars[i].get() == "on" else False for i in range(self.num_of_acts)]
            actnames=[self.my_frame.choices[i].cget("text") for i in range(len(self.my_frame.choices))]

            # Making sure to save all of the activities:
            data_sheet.delete_cols(1,1)
            data_sheet["A1"].value = self.num_of_acts
            for i in range(2, self.num_of_acts + 2):
                data_sheet[f"A{i}"].value = actnames[i]
            data_book.save("past_activity_data.xlsx")


            act_dict = []
            #Sort out the categories
            categs = ["All" if self.my_frame.checkboxboy[i].get() == self.my_frame.checkboxgirl[i].get() else "JustBoy" if self.my_frame.checkboxboy[i].get() == "on" else "JustGirl" for i in range(self.num_of_acts)]
            
            categs = [categs[i]+"Simul" if self.my_frame.checkboxsimul[i].get() == "on" else categs[i] for i in range(self.num_of_acts)]

            categs = [categs[i]+"Double" if self.my_frame.checkboxSpan[i].get() == "on" else categs[i] for i in range(self.num_of_acts)]

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

                # Sending Produced Schedule to Downloads
                output_path=Path.home()/'Downloads'/f'Generated_Schedule_{datetime.now().month}_{datetime.now().day}_{datetime.now().year}.xlsx'
                shutil.copyfile('new_sheet.xlsx',output_path)
                print(f"Here is how similar it is to previous schedules (140 meaning identical): {compareAgainstPrevious(proposed_sched=result, previous_scheds=previous_schedules)}")
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
            
            def compareAgainstPrevious(proposed_sched: list, previous_scheds: list):
                # Currently only works for whole day schedules.
                same_acts = 0
                same_times = 0
                total_score = 0

                proposed = np.array(proposed_sched).transpose()

                for previous in previous_scheds:
                    for new_row, old_row,  in zip(proposed, np.array(previous).transpose()):
                        new_row = [act.replace(" ", "") for act in new_row] # Stripped all spaces in case of random mistakes space-wise
                        old_row = [act.replace(" ", "") for act in old_row]
                        same_times = sum([2 for new_act, old_act in zip(new_row, old_row) if new_act == old_act])
                        same_acts = sum([1 for old_act in old_row if old_act in new_row])
                    total_score = same_times + same_acts if same_acts + same_times > total_score else total_score


                return total_score


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
        self.choicemorning.grid(row=7, column=4,
                          padx=5, pady=5,
                          sticky="ew")
        
        self.choiceafternoon = ctk.CTkCheckBox(self,
                            text="Afternoon",
                            variable=self.checkboxafternoon,
                            onvalue="on",
                            offvalue="off",
                            command=reset_checkboxes_afternoon)                             
        self.choiceafternoon.grid(row=7, column=5,
                          padx=5, pady=5,
                          sticky="ew")

        self.choicewholeday = ctk.CTkCheckBox(self,
                            text="Whole Day",
                            variable=self.checkboxwholeday,
                            onvalue="on",
                            offvalue="off",
                            command=reset_checkboxes_wholeday)                               
        self.choicewholeday.grid(row=7, column=6,
                          padx=5, pady=5,
                          sticky="ew")


        #Make Spreadsheet Button
        self.testButton=ctk.CTkButton(self,text="Deselect All Activities", text_color="#000000", corner_radius=32, fg_color="#FFFFFF",hover_color="#FFFFFF",command=deselectall)
        self.testButton.grid(row=1,column=4,columnspan=3,padx=10,pady=10,sticky="ew")

        self.saveButton=ctk.CTkButton(self,text="Save Previous", text_color="#000000", corner_radius=32, fg_color="#77DD77",hover_color="#B9FEB9",command=deselectall)
        self.saveButton.grid(row=1,column=3,columnspan=1,padx=5,pady=5,sticky="ew")

        self.eraseButton=ctk.CTkButton(self,text="Erase Memory", text_color="#000000", corner_radius=32, fg_color="#FF6961",hover_color="#f69697",command=deselectall)
        self.eraseButton.grid(row=1,column=7,columnspan=1,padx=5,pady=5,sticky="ew")

        self.testButton=ctk.CTkButton(self,text="Select Duration of Schedule", text_color="#000000", corner_radius=32, fg_color="#FFFFFF",hover_color="#FFFFFF",command=testrun)
        self.testButton.grid(row=6,column=4,columnspan=3,padx=10,pady=10,sticky="ew")

        self.generateResultsButton=ctk.CTkButton(self,text="Generate Schedule", corner_radius=32, fg_color="#0000FF",hover_color="#33BFFF",command=run)
        self.generateResultsButton.grid(row=8,column=4,columnspan=3,padx=10,pady=10,sticky="ew")

        self.testButton=ctk.CTkButton(self,text=errormessage, text_color="#000000", corner_radius=32, fg_color="#FFFFFF",hover_color="#FFFFFF",command=testrun)
        self.testButton.grid(row=9,column=4,columnspan=3,padx=10,pady=10,sticky="ew")
      
        #my_image=ctk.CTkImage(Image.open('images/CampIHCLogo2.png'),size=(150,150))
        #my_image_label=ctk.CTkLabel(self, image=my_image, width=200, height=200)
        #my_image_label.grid(row=7,column=9,columnspan=1,padx=10,pady=10,sticky="ew")
        #my_image_label.grid_propagate(False)

if __name__ == "__main__":
    app = App()
    # Runs the app
    app.mainloop()