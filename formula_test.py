import math
import time
import pyautogui as pag


def calculate_angle_and_length(x_11, y_11):
    angle = math.atan2(y_11, x_11)
    angle_degrees = math.degrees(angle)
    angle_round = round(angle_degrees, 2)
    length = math.sqrt(x_11**2 + y_11**2)
    length_round = round(length)
    return angle_round, length_round


if __name__ == '__main__':
    act = True
    while act:
        ttr = str(input("start y/n?: "))
        if ttr == 'y':
            x = int(input("how much iterations?: "))
            y = float(input("interval sec.?: "))
            count = 1
            while count <= x:
                count += 1
                s = str(pag.position())
                s_segment = s[5:].replace('=', '').replace('(', '').replace(')', '').replace('x', '').replace('y', '')
                s_split = s_segment.split(',')
                s_x = int(s_split[0])
                s_y = int(s_split[1])
                mathim = str(calculate_angle_and_length(s_x, s_y)).replace('(', '').replace(')', '')
                mathim_split = mathim.split(',')
                degr = str(mathim_split[0])
                le = str(mathim_split[1])
                le_1 = le.replace(' ', '')
                print(f'x:{s_x} y:{s_y}')
                print(f'градус:{degr} длина:{le_1}')
                time.sleep(y)
        else:
            act = False

    print('close')
    exit()
