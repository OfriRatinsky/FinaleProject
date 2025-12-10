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

    def run_exercise(self, name, hand=''):
        s.success_exercise = False
        print("TRAINING: Exercise ", name, " start")
        say(name+hand)
        # time.sleep(3)  # Delay the robot movement after the audio is played
        s.req_exercise = name
        while s.req_exercise == name:
            time.sleep(0.001)  # Prevents the MP to stuck
        if s.success_exercise:
            say(self.random_encouragement())
        print("TRAINING: Exercise ", name, " done")
        time.sleep(1)

    def random_encouragement(self):
        enco = ["well done", "very good", "excellent"]
        return random.choice(enco)


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