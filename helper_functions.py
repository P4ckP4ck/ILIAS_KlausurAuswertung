import os
import numpy as np
import pandas as pd


class ResultDict:
    # Eine Klasse zur einfacheren Handhabung der ILIAS-Daten. Sie soll die spätere Umwandlung in ein Pandas DataFrame
    # vereinfachen
    def __init__(self):
        self.data = dict()
        self.std_id = 0

    def _check_student_id(self, student_id):
        # Überprüfung, ob eine neue Zeile eingelesen wurde
        if not self.std_id == student_id:
            self.std_id = student_id
            self._check_dict_consistency()

    def _check_dict_consistency(self):
        # Die Funktion stellt sicher, dass die Längen der einzelnen Reihen während der generierung des DataFrames
        # gleich bleiben! Im Falle einer Abweichung (eine Zeile ist leer) wird eine -9999 eingefügt
        for question in self.data.keys():
            length_check = [len(self.data[question][x]) for x in self.data[question].keys()]
            question_list = list(self.data[question].keys())
            if not length_check.count(length_check[0]) == len(length_check):
                idx = length_check.index(min(length_check))
                self.data[question][question_list[idx]].append("-9999")

    def append(self, current_question, input, unique_id, student_id):
        # Funktion zum korrekten anhängen in das DataFrame
        self._check_student_id(student_id)
        if input.Question is np.nan:
            input.Question = unique_id
        if input.Answer is np.nan:
            input.Answer = -9999
        if current_question not in self.data.keys():
            # checks if the current question is already in the dict, and creates a new one if not
            self.data[current_question] = dict()
        if input.Question not in self.data[current_question].keys():
            self.data[current_question][input.Question] = [input.Answer]
        else:
            self.data[current_question][input.Question].append(input.Answer)

    def save(self):
        self._check_dict_consistency()
        dir = "./answers_per_question/"
        if not "answers_per_question" in os.listdir():
            os.makedirs(dir)
        for key in self.data.keys():
            pd.DataFrame(self.data[key]).to_csv(dir + key + ".csv")