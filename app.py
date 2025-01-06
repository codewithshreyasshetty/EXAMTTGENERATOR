import pandas as pd
from ortools.sat.python import cp_model

# Define data structures
class Exam:
    def __init__(self, subject, date, time_slot, room, invigilator):
        self.subject = subject
        self.date = date
        self.time_slot = time_slot
        self.room = room
        self.invigilator = invigilator

class Timetable:
    def __init__(self):
        self.exams = []

    def add_exam(self, exam):
        self.exams.append(exam)

    def display(self):
        for exam in self.exams:
            print(f"{exam.subject} | {exam.date} | {exam.time_slot} | Room: {exam.room} | Invigilator: {exam.invigilator}")

# Input Details
def get_input_data():
    subjects = input("Enter subjects (comma-separated): ").split(",")
    students_per_subject = list(map(int, input("Enter number of students enrolled in each subject (comma-separated): ").split(",")))
    dates = input("Enter available exam dates (comma-separated): ").split(",")
    time_slots = input("Enter available time slots (comma-separated): ").split(",")
    rooms = input("Enter room capacities (comma-separated): ").split(",")
    invigilators = input("Enter invigilators (comma-separated): ").split(",")
    
    return subjects, students_per_subject, dates, time_slots, rooms, invigilators

# Conflict Detection and Scheduling
def schedule_exams(subjects, students_per_subject, dates, time_slots, rooms, invigilators):
    model = cp_model.CpModel()
    
    # Create variables for each exam
    exam_vars = {}
    
    for i, subject in enumerate(subjects):
        for date in dates:
            for time in time_slots:
                for room in rooms:
                    var_name = f"{subject}_{date}_{time}_{room}"
                    exam_vars[var_name] = model.NewBoolVar(var_name)

    # Constraints: No overlapping exams for students
    # Add your constraints here based on student enrollments and other requirements
    
    # Objective: Minimize the number of exams scheduled (or other criteria)
    
    # Solve the model
    solver = cp_model.CpSolver()
    status = solver.Solve(model)
    
    timetable = Timetable()
    
    if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
        for var_name in exam_vars:
            if solver.Value(exam_vars[var_name]) == 1:
                subject, date, time, room = var_name.split('_')
                invigilator = invigilators[0]  # Simplified; assign based on availability logic
                timetable.add_exam(Exam(subject.strip(), date.strip(), time.strip(), room.strip(), invigilator))
    
    return timetable

# Export to Excel
def export_to_excel(timetable):
    df = pd.DataFrame([vars(exam) for exam in timetable.exams])
    df.to_excel("examination_timetable.xlsx", index=False)

# Main function to run the program
def main():
    subjects, students_per_subject, dates, time_slots, rooms, invigilators = get_input_data()
    
    timetable = schedule_exams(subjects, students_per_subject, dates, time_slots, rooms, invigilators)
    
    print("\nGenerated Examination Timetable:")
    timetable.display()
    
    export_to_excel(timetable)
    
if __name__ == "__main__":
    main()
