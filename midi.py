import rtmidi_python as rtmidi
import re


def midi():
    midi_in = rtmidi.MidiIn()
    midi_in.open_port(0)
    message, delta_time = midi_in.get_message()
    while True:
        message, delta_time = midi_in.get_message()
        if message:
            knob_num = str(message[1])
            num = 'in.'
            knob_value = str(message[2])
            val = 'val.'
            midi_val = val
            print(knob_value)
            return (midi_val)
        

if __name__ == '__main__':
    while True:
        midi()
