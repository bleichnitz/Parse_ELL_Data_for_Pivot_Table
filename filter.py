def filter_for_student(student,list):
    """
    Take one student and compare the student number aginst a list. If the student is on the
    list, return return the student ROW data, if not on the list return FALSE
    :param student:
    :param list:
    :return:
    """
    content = False
    student_data = []
    for i in range(0,len(list)):
        if student == list[i][0]:
            student_data.append(list[i])
            content = student_data
            return content
    return False