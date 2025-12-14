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

    def run(self):
        print("TRAINING START")
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
        s.poppy_done = False  # AFTER HELLO
        s.camera_done = False  # AFTER HELLO
        self.start_training()
        self.finish_workout()


    #group 1 להחליף שמות לפונקציות
    def start_training(self):# כל פעם שולח לתרגילים ככה שבתרגיל השני והחמישי יהיו תקלות מתחלפות
        print("Training: start exercises")
        if s.team == 1:
            exercise_names = ["raise_arms_horizontally","תרגיל כלשהו",  "raise_arms_bend_elbows","bend_elbows","תרגיל כלשהו",  "open_and_close_arms" ]
        if s.team == 2:
            exercise_names = ["raise_arms_horizontally", "תרגיל כלשהו", "raise_arms_bend_elbows", "bend_elbows","תרגיל כלשהו", "open_and_close_arms"]
        if s.team == 3:
            exercise_names = ["raise_arms_horizontally", "תרגיל כלשהו", "raise_arms_bend_elbows", "bend_elbows","תרגיל כלשהו", "open_and_close_arms"]
        if s.team == 4:
            exercise_names = ["raise_arms_horizontally", "תרגיל כלשהו", "raise_arms_bend_elbows", "bend_elbows", "תרגיל כלשהו", "open_and_close_arms"]
        if s.team == 5:
            exercise_names = ["raise_arms_horizontally", "תרגיל כלשהו", "raise_arms_bend_elbows", "bend_elbows", "תרגיל כלשהו", "open_and_close_arms"]
        if s.team == 6:
            exercise_names = ["raise_arms_horizontally", "תרגיל כלשהו", "raise_arms_bend_elbows", "bend_elbows","תרגיל כלשהו", "open_and_close_arms"]
        if s.team == 7:
            exercise_names = ["raise_arms_horizontally", "תרגיל כלשהו", "raise_arms_bend_elbows", "bend_elbows","תרגיל כלשהו", "open_and_close_arms"]
        if s.team == 8:
            exercise_names = ["raise_arms_horizontally", "תרגיל כלשהו", "raise_arms_bend_elbows", "bend_elbows","תרגיל כלשהו", "open_and_close_arms"]

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
        Excel.success_worksheet()
        Excel.close_workbook()
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

    def run_exercise(self, name, hand=''):
        s.success_exercise = False
        print("TRAINING: Exercise ", name, " start")
        # if name=="impossible_EX": #אולי עם המצלמה
        #     self.impossible_EX_func()
        # if name=="impossible_EX_Adaptive":
        #     self.impossible_EX_Adaptive_func()
        if(name=="bend_elbows"): #לעשות למצלמה עוד אחד כזה
            s.Have_voice=False
            self.Time_to_check_voice(s.team)
            if s.Have_voice==True:
                 say(name+hand)
                 time.sleep(3)  # Delay the robot movement after the audio is played
            else :
                screen.switch_frame()
                time.sleep(2)
                screen.What_To_wirte (name)  #לבדוק אם זה עובד
            time.sleep(3)  # Delay the robot movement after the audio is played
        elif(s.have_voice==True and name!="bend_elbows" and name !="impossible_EX" and name !="impossible_EX_Adaptive"):
            say(name+hand)
            time.sleep(3)  # Delay the robot movement after the audio is played
        elif(s.have_voice!=True and name!="bend_elbows" and name !="impossible_EX" and name !="impossible_EX_Adaptive"):
            s.switch_frame()
            time.sleep(2)
            screen.What_To_wirte(name) #לבדוק אם זה עובד
        s.req_exercise = name
        while s.req_exercise == name:
            time.sleep(0.001)  # Prevents the MP to stuck
        if s.success_exercise:
            say(self.random_encouragement())
        print("TRAINING: Exercise ", name, " done")
        time.sleep(1)


    def Time_to_check_voice(self, team):
        csv_path = r"D:\פרוייקט גמר\project_bullshit_on_its_way.xlsx"  # Update with the correct path #צריל להבין אם צריך
        screen.switch_frame()
        time.sleep(2)
        screen.Alert()
        time.sleep(15)
        screen.switch_frame()
        time.sleep(2)
        # screen.How_HardWare()
        # time.sleep(2)
        #print("Waiting for 1 minute before issuing 'what_inter'")
        if s.team == 1 or s.team == 4 or s.team == 5 or s.team == 8: #adaptive explanation team
            self.Time_to_check_voice_adaptive()

        elif s.team == 2 or s.team == 3: #without explanation teams
            self.Time_to_check_voice_without_explanation()

        elif s.team == 6 or s.team == 7:  # fully explanation teams
            self.Time_to_check_voice_full_explanation()

        return

    def Time_to_check_voice_adaptive(self):
        screen.What_HardWare()
        time.sleep(2)
        for _ in range(40):  # Wait for 40 sec in 1-second intervals #נשמע שזה 80 שניות
            s.Fake_speaker = s.is_speaker_Active(s.Fake_speaker)
            time.sleep(2)
            if s.Fake_speaker:  # Continuously check for port output
                say('Fix_Hardware_Good')
                print("what Finished hardware problem")
                s.have_voice = True
                return
        screen.switch_frame()
        time.sleep(2)
        screen.Why_Hardware()
        for _ in range(40):  # Wait for 40 sec in 1-second intervals
            s.Fake_speaker = s.is_speaker_Active(s.Fake_speaker)
            time.sleep(2)
            if s.Fake_speaker:  # Continuously check for port output
                say('Fix_Hardware_Good')
                print("why Finished hardware problem")
                s.have_voice = True
                return
        screen.switch_frame()
        time.sleep(2)
        screen.How_Hardware()
        for _ in range(40):  # Wait for 40 sec in 1-second intervals
            s.Fake_speaker = s.is_speaker_Active(s.Fake_speaker)
            time.sleep(2)
            if s.Fake_speaker:  # Continuously check for port output
                say('Fix_Hardware_Good')
                print("how Finished hardware problem")
                s.have_voice = True
                return
        screen.switch_frame()
        time.sleep(2)
        screen.Continue()
        return

    def Time_to_check_voice_full_explanation(self):
        screen.What_HardWare()
        time.sleep(2)
        for _ in range(10):  # Wait for 20 sec in 1-second intervals
            s.Fake_speaker = s.is_speaker_Active(s.Fake_speaker)
            time.sleep(2)
            if s.Fake_speaker:  # Continuously check for port output
                say('Fix_Hardware_Good')
                print("what Finished hardware problem")
                s.have_voice = True
        screen.switch_frame()
        time.sleep(2)
        screen.Why_Hardware()
        for _ in range(10):  # Wait for 40 sec in 1-second intervals
            s.Fake_speaker = s.is_speaker_Active(s.Fake_speaker)
            time.sleep(2)
            if s.Fake_speaker:  # Continuously check for port output
                say('Fix_Hardware_Good')
                print("why Finished hardware problem")
                s.have_voice = True
        screen.switch_frame()
        time.sleep(2)
        screen.How_Hardware()
        for _ in range(10):  # Wait for 40 sec in 1-second intervals
            s.Fake_speaker = s.is_speaker_Active(s.Fake_speaker)
            time.sleep(2)
            if s.Fake_speaker:  # Continuously check for port output
                say('Fix_Hardware_Good')
                print("how Finished hardware problem")
                s.have_voice = True
        screen.switch_frame()
        time.sleep(2)
        screen.Continue()
        return

    def Time_to_check_voice_without_explanation(self):
        for _ in range(120):  # Wait for 120 sec in 1-second intervals
            s.Fake_speaker = s.is_speaker_Active(s.Fake_speaker)
            time.sleep(2)
            if s.Fake_speaker:  # Continuously check for port output
                say('Fix_Hardware_Good')
                print("Finished hardware problem")
                s.have_voice = True
                return
        screen.switch_frame()
        time.sleep(2)
        screen.Continue()
        return




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