import argparse
import os
import subprocess
from platform import system as system_name  # Returns the system/OS name

import telebot
import yaml

TOKEN = ''
DIALOG_ID = 0


def ping(host):
    parameters = "-n 1" if system_name().lower() == "windows" else "-c 1"
    return subprocess.call('ping ' + parameters + ' ' + host,
                           stdout=subprocess.PIPE,
                           stderr=subprocess.PIPE,
                           shell=True) == 0


if __name__ == '__main__':
    arguments = argparse.ArgumentParser()
    arguments.add_argument('--ip', help='IP of the computer', required=True)
    arguments = arguments.parse_args()

    last_state = {}
    if os.path.exists('last_state.yaml'):
        with open('last_state.yaml') as cache_file:
            last_state = yaml.load(cache_file)

    current_state = ping(arguments.ip)
    print arguments.ip, 'ping', current_state

    if current_state != last_state.get(arguments.ip):
        bot = telebot.TeleBot(TOKEN)
        bot.send_message(DIALOG_ID, '{} is {}'.format(arguments.ip, 'ok' if current_state else 'frozen'))
    last_state[arguments.ip] = current_state

    with open('last_state.yaml', 'wt') as cache_file:
        yaml.dump(last_state, cache_file)
