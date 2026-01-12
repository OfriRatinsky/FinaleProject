import threading
import time
import Settings as s
import Excel
import random
from Audio import say
import Screen as screen #הוספתי
import Camera as cam #הוספתי
import os #הוספתי



class Training(threading.Thread):
    def __init__(self):
        threading.Thread.__init__(self)
        self.experiment_rows = []

    def run(self):
        print("TRAINING START")#
        self.run_exercise("hello_waving")
        print("Training: start waving")
        while not s.waved:
            time.sleep(0.00000001)  # Prevents the MP to stuck
            continue
        s.waved = False # set as False again for future
        if not s.calibration:
            print("Training: Calibration")
            s.camera.init_position()
            while not s.calibration:
                time.sleep(0.00000001)
                continue
        time.sleep(3)
        say('lets start')
        time.sleep(2.5)
        print("Training: finish waving")
        self.warm_up()
        time.sleep(8)
        print("Training: finish warmup")
        s.poppy_done = False  # AFTER HELLO
        s.camera_done = False  # AFTER HELLO
        self.start_training()
        self.finish_workout()

    def warm_up(self):
        say('start_warm_up')
        print("Training: start exercises")
        # TODO - adding random choice of exercises.
        exercise_names = ["open_and_close_arms_90","open_and_close_arms"]
        for e in exercise_names:
            time.sleep(1) # wait between exercises
            self.run_exercise(e)
            while (not s.poppy_done) or (not s.camera_done):
                print("not done")
                time.sleep(1)
        say('end_warm_up')
        time.sleep(1)

    #group 1 להחליף שמות לפונקציות
    def start_training(self):# כל פעם שולח לתרגילים ככה שבתרגיל השני והחמישי יהיו תקלות מתחלפות
        print("Training: start exercises")
        if s.team == 1:
            exercise_names = ["bend_elbows"]#, "raise_arms_forward","open_and_close_arms_90"]

           #exercise_names = ["raise_arms_horizontally","bend_elbows",  "raise_arms_bend_elbows","break","open_and_close_arms_90","raise_arms_forward",  "open_and_close_arms" ]
        if s.team == 2:
            exercise_names = ["raise_arms_forward"]
            #exercise_names = ["raise_arms_horizontally", "raise_arms_forward", "raise_arms_bend_elbows","break", "open_and_close_arms_90","bend_elbows", "open_and_close_arms"]
        if s.team == 3:
            exercise_names = ["raise_arms_horizontally", "bend_elbows", "raise_arms_bend_elbows","break", "open_and_close_arms_90","raise_arms_forward", "open_and_close_arms"]
        if s.team == 4:
            exercise_names = ["raise_arms_horizontally", "raise_arms_forward", "raise_arms_bend_elbows","break", "open_and_close_arms_90", "bend_elbows", "open_and_close_arms"]
        if s.team == 5:
            exercise_names = ["raise_arms_horizontally", "bend_elbows", "raise_arms_bend_elbows","break", "open_and_close_arms_90", "raise_arms_forward", "open_and_close_arms"]
        if s.team == 6:
            exercise_names = ["raise_arms_horizontally", "raise_arms_forward", "raise_arms_bend_elbows","break", "open_and_close_arms_90", "bend_elbows", "open_and_close_arms"]
        if s.team == 7:
            exercise_names = ["raise_arms_horizontally", "bend_elbows", "raise_arms_bend_elbows","break", "open_and_close_arms_90","raise_arms_forward", "open_and_close_arms"]
        if s.team == 8:
            exercise_names = ["raise_arms_horizontally", "raise_arms_forward", "raise_arms_bend_elbows","break", "open_and_close_arms_90", "bend_elbows", "open_and_close_arms"]

        for e in exercise_names:
            time.sleep(2) # wait between exercises
            self.run_exercise(e)
            while (not s.poppy_done) or (not s.camera_done):
                print("not done")
                time.sleep(1)

    def training_session(self):
        print("Training: start exercises")
        # TODO - adding random choice of exercises.
        #exercise_names = ["raise_arms_horizontally", "bend_elbows", "raise_arms_bend_elbows", "open_and_close_arms",
        #                  "open_and_close_arms_90", "raise_arms_forward"]
        #exercise_names = ["raise_arms_horizontally"]
        exercise_names = ["claps"]
        for e in exercise_names:
            time.sleep(2) # wait between exercises
            self.run_exercise(e)
            while (not s.poppy_done) or (not s.camera_done):
                print("not done")
                time.sleep(1)

    def finish_workout(self):
        say('goodbye')
        s.finish_workout = True
        s.robot.loose_motors()
        Excel.success_worksheet()
        Excel.close_workbook()
        print("ROWS TO SAVE =", len(self.experiment_rows))
        if self.experiment_rows:
            print("FIRST ROW =", self.experiment_rows[0])

        excel_path = r"C:\Users\owner\Desktop\repair_time.xlsx"
        Excel.append_rows_to_excel(excel_path, self.experiment_rows)
        time.sleep(10)
        s.screen.quit()
        print("TRAINING DONE")

    # def run_exercise(self, name, hand=''):
    #     s.success_exercise = False
    #     print("TRAINING: Exercise ", name, " start")
    #     say(name+hand)
    #     # time.sleep(3)  # Delay the robot movement after the audio is played
    #     s.req_exercise = name
    #     while s.req_exercise == name:
    #         time.sleep(0.001)  # Prevents the MP to stuck
    #     if s.success_exercise:
    #         say(self.random_encouragement())
    #     print("TRAINING: Exercise ", name, " done")
    #     time.sleep(1)

    def What_To_write (self,name):####לשנות תרגילים
        if(name=='raise_arms_horizontally'):#######
            s.screen.switch_frame(screen.raise_arms_horizontally)
        if(name=='open_and_close_arms_90'):#########
            s.screen.switch_frame(screen.open_and_close_arms_90)
        if(name=='raise_arms_bend_elbows'):
            s.screen.switch_frame(screen.raise_arms_bend_elbows)
        if(name=='open_and_close_arms'):
            s.screen.switch_frame(screen.open_and_close_arms)
        if(name=='raise_arms_forward'):
            s.screen.switch_frame(screen.raise_arms_forward)

    def run_exercise(self, name, hand=''):
        s.success_exercise = False
        print("TRAINING: Exercise ", name, " start")
        if(name == "break"): #Questionnaire completion
            input("Press Enter to continue...")
            return
        if(name== "raise_arms_forward"):
            s.camera_not_recognize = True
            if s.have_voice==True:
                say(name+hand)
            if s.have_voice!=True:
                s.screen.switch_frame(screen.raise_arms_forward) 
            time.sleep(3)  # Delay the robot movement after the audio is played
            s.req_exercise = name
            self.Time_to_check_camera(s.team)
        if(name=="bend_elbows"): #לעשות למצלמה עוד אחד כזה
            s.have_voice=False
            s.screen.switch_frame(screen.bend_elbows)
            time.sleep(1)
            s.req_exercise = name
            self.Time_to_check_voice(s.team)
            
        elif(s.have_voice==True and name!="bend_elbows" and name !="raise_arms_forward"):
            say(name+hand)
            time.sleep(3)  # Delay the robot movement after the audio is played
            s.req_exercise = name
        elif(s.have_voice!=True and name!="bend_elbows" and name !="raise_arms_forward"):
            self.What_To_write(name) #לבדוק אם זה עובד
            time.sleep(2)
            s.req_exercise = name
        while s.req_exercise == name:
            time.sleep(0.001)  # Prevents the MP to stuck
       # if s.success_exercise:
        print("TRAINING: Exercise ", name, " done")
        time.sleep(1)


    def Time_to_check_voice(self, team):
        csv_path = r"D:\חיבורמקול.docx"  # Update with the correct path #צריל להבין אם צריך
        # Start with the Alert frame
        time.sleep(3)
        if s.team == 1 or s.team == 4 or s.team == 5 or s.team == 8: #adaptive explanation team
            s.screen.switch_frame(screen.Alert)
            time.sleep(1)
            self.Time_to_check_voice_adaptive(csv_path)

        elif s.team == 2 or s.team == 3: #without explanation teams
            self.Time_to_check_voice_without_explanation(csv_path)

        elif s.team == 6 or s.team == 7:  # fully explanation teams
            s.screen.switch_frame(screen.Alert)
            time.sleep(1)
            self.Time_to_check_voice_full_explanation(csv_path)

        return

    def Time_to_check_voice_adaptive(self, csv_path):
        start = time.perf_counter()
        hardware_stages = [
            (screen.What_Hardware, "what Finished hardware problem"),
            (screen.Why_Hardware, "why Finished hardware problem"),
            (screen.How_Hardware, "how Finished hardware problem"),
            (screen.Continue, "Finished hardware check, no solution found"),
        ]
        for frame, message in hardware_stages[:-1]:  # Exclude the "Continue" stage for now
            s.screen.switch_frame(frame)
            time.sleep(2)
            print(f"Checking for speaker activity during '{frame.__name__}'")
            
            for _ in range(40):  # Check for 40 seconds in 1-second intervals
                s.Fake_speaker = self.is_speaker_Active(csv_path)
                time.sleep(1)
                
                if s.Fake_speaker:  # If speaker is active
                    s.have_voice = True
                    print(message)
                    say("Fix_Hardware_Good")
                    s.screen.switch_frame(screen.EyesPage)
                    end = time.perf_counter()
                    seconds = round(end - start)
                    print(seconds)
                    self.log_repair("hardware", "adaptive", seconds, message.split()[0])
                    return
                
        s.screen.switch_frame(hardware_stages[-1][0])  # "Continue" frame
        print(hardware_stages[-1][1])
        time.sleep(2)

    def Time_to_check_voice_full_explanation(self, csv_path):
        start = time.perf_counter()
        hardware_stages = [
            (screen.What_Hardware, "what Finished hardware problem"),
            (screen.Why_Hardware, "why Finished hardware problem"),
            (screen.How_Hardware, "how Finished hardware problem"),
            (screen.Continue, "Finished hardware check, no solution found"),
        ]
        for frame, message in hardware_stages[:-1]:  # Exclude the "Continue" stage for now
            s.screen.switch_frame(frame)
            time.sleep(2)
            print(f"Checking for speaker activity during '{frame.__name__}'")
            
            for _ in range(15):  # Check for 15 seconds in 1-second intervals
                s.Fake_speaker = self.is_speaker_Active(csv_path)
                time.sleep(1)
                
                if s.Fake_speaker:  # If speaker is active
                    s.have_voice = True
                    print(message)
                    say("Fix_Hardware_Good")
                    s.screen.switch_frame(screen.EyesPage)
                    end = time.perf_counter()
                    seconds = round(end - start)
                    print(seconds)
                    self.log_repair("hardware", "full", seconds)
                    return

        if s.have_voice!=True:
            s.screen.switch_frame(hardware_stages[-1][0])  # "Continue" frame
            print(hardware_stages[-1][1])
            time.sleep(2)
        

    def Time_to_check_voice_without_explanation(self, csv_path):
        start = time.perf_counter()
        for _ in range(120):  # Wait for 120 sec in 1-second intervals
            s.Fake_speaker = self.is_speaker_Active(csv_path)
            time.sleep(1)
            if s.Fake_speaker:  # Continuously check for port output
                say('Fix_Hardware_Good')
                print("Finished hardware problem")
                s.have_voice = True
                end = time.perf_counter()
                seconds = round(end - start)
                print(seconds)
                self.log_repair("hardware", "without", seconds)

                return
        s.screen.switch_frame(screen.Continue)
        print("Finished hardware check, no solution found")
        time.sleep(2)
        return
    
    def is_speaker_Active(self, path):
        try:
        # Check if the file exists
         if os.path.exists(path):
            #pd.read_excel(path)  # Attempt to import the file
            print("File imported successfully!")
            s.screen.switch_frame(screen.EyesPage)
            return True
         else:
            print(f"File does not exist at: {path}")
            return False
        except Exception as e:
         print(f"Error while trying to import the file: {e}")
         s.screen.switch_frame(screen.EyesPage)
        return True      
    

    def Time_to_check_camera(self, team):
        time.sleep(3) #15
        if s.team == 2 or s.team == 3 or s.team == 6 or s.team == 7: #adaptive explanation team
            s.screen.switch_frame(screen.Alert)
            time.sleep(1)
            self.Time_to_check_camera_adaptive()

        elif s.team == 1 or s.team == 4: #without explanation teams
            self.Time_to_check_camera_without_explanation()

        elif s.team == 5 or s.team == 8:  # fully explanation teams
            s.screen.switch_frame(screen.Alert)
            time.sleep(1)
            self.Time_to_check_camera_full_explanation()

        return

    def Time_to_check_camera_adaptive(self):
        start = time.perf_counter()
        inter_stages = [
            (screen.What_inter,"what Finished inter problem"),
            (screen.Why_inter,"why Finished inter problem"),
            (screen.How_inter, "how Finished inter problem"),
            (screen.continue_inter, "Finished inter check, no solution found"),
        ]
        for frame, message in inter_stages[:-1]:  # Exclude the "Continue" stage for now
            s.screen.switch_frame(frame)
            if s.have_voice==True:
                say(frame.__name__)#פה שניתי!!!
            time.sleep(2)
            
            for _ in range(40):  # Check for 40 seconds in 1-second intervals
                #self.is_camera_Active()
                time.sleep(1)
                
                if s.Fake_camera:  # If camera is active
                    s.camera_not_recognize = False
                    print(message)
                    say("Fix_inter_good")
                    s.screen.switch_frame(screen.EyesPage)

                    end = time.perf_counter()
                    seconds = round(end - start)
                    print(seconds)
                    self.log_repair("interaction", "adaptive", seconds,message.split()[0])

                    return 
                
        if s.have_voice == True:
            say("continue_inter")
        else:
            s.screen.switch_frame(screen.continue_inter)
        print(inter_stages[-1][1])
        time.sleep(2)

    def Time_to_check_camera_full_explanation(self):
        start = time.perf_counter()
        inter_stages = [
            (screen.What_inter, "what Finished inter problem"),
            (screen.Why_inter, "why Finished inter problem"),
            (screen.How_inter, "how Finished inter problem"),
            (screen.continue_inter, "Finished inter check, no solution found"),
        ]
        for frame, message in inter_stages[:-1]:  # Exclude the "Continue" stage for now
            s.screen.switch_frame(frame)
            if s.have_voice==True:
                say(frame.__name__)
            time.sleep(2)
            
            for _ in range(15):  # Check for 15 seconds in 1-second intervals
                time.sleep(1)
                
                if s.Fake_camera:  # If speaker is active
                    s.camera_not_recognize = False
                    print(message)
                    say("Fix_inter_good")
                    s.screen.switch_frame(screen.EyesPage)

                    end = time.perf_counter()
                    seconds = round(end - start)
                    print(seconds)
                    self.log_repair("interaction", "full", seconds)
                    return
        if s.camera_not_recognize == True:
            if s.have_voice == True:
                say("continue_inter")
            else:
                s.screen.switch_frame(screen.continue_inter)
            print(inter_stages[-1][1])
            time.sleep(2)
        

    def Time_to_check_camera_without_explanation(self):
        start = time.perf_counter()
        for _ in range(120):  # Wait for 120 sec in 1-second intervals
            time.sleep(1)
            if s.Fake_camera:  # Continuously check for port output
                if s.have_voice == True:
                    say('Fix_inter_good')
                    s.screen.switch_frame(screen.EyesPage)
                else:
                    s.screen.switch_frame(screen.Fix_inter_good)
                print("Finished inter problem")
                s.camera_not_recognize = False
                end = time.perf_counter()
                seconds = round(end - start)
                print(seconds)
                self.log_repair("interaction","without",seconds)

                return
            
        if s.have_voice == True:
            say("continue_inter")
        else:
            s.screen.switch_frame(screen.continue_inter)
        print("Finished hardware check, no solution found")
        time.sleep(2)
        return

    def log_repair(self, failure_type, feedback_type, repair_time_s, feedback_stage=""):
        self.experiment_rows.append({
            "Participant Number": s.participant_number,
            "group number": s.team,
            "hardware/inter": failure_type,  # "hardware" / "interaction"
            "Feedback Type": feedback_type,
            "Repair Time (s)": repair_time_s,
            "Feedback Stage (if adaptive)": feedback_stage or ""
        })



if __name__ == "__main__":
    # Create all components
    from Camera import Camera
    from Poppy import Poppy

    s.camera = Camera()
    s.robot = Poppy()
    language = 'Hebrew'
    gender = 'Male'
    s.audio_path = 'audio files/' + language + '/' + gender + '/'
    s.finish_workout = False
    s.rep = 8 #todo change to 8
    s.req_exercise = ""
    s.robot_count = True

    # Adaptation variables
    s.adaptive = True
    s.corrective_feedback = True
    s.one_hand = False
    s.robot_rep = 0
    if s.adaptive:
        s.adaptation_model_name = 'performance_evaluation_model'
        s.performance_class = {}
    s.camera.start()
    s.robot.start()

    t = Training()
    t.run_exercise("open_and_close_arms_90")