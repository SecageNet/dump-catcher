import click
from os import path, mkdir
import shutil
import subprocess
import glob
from concurrent.futures import ThreadPoolExecutor
import sys
import locale
from json import load

TEMP_DIR = path.join(path.expanduser("~"),"AppData","Local","Temp","SystemInformationReports")

def load_phrases():
    with open('phrases.json', 'r', encoding='utf-8') as f:
        return load(f)

phrases = load_phrases()

turkeyName = ["tr_TR", "Turkish_Turkey", "Turkish_Türkiye"]

lang = locale.getlocale()[0]
lang = 'tr' if lang in turkeyName else 'en'

def phrase(key):
    return phrases[key][lang]

REPORT_ZIP = phrase('report_filename')

def run_subprocess(command):
    subprocess.run(command, shell=True, check=True)

def custom_pause(message):
    click.echo(message)
    while True:
        char = click.getchar()
        if char.lower() == 'q':
            click.echo(phrase('goodbye'))
            sys.exit(0)
        else:
            break

def compress_and_clean(temp_dir, output_name):
    print(phrase('compressing'))
    output_path = path.join(path.expanduser("~"),"Desktop",output_name.replace('.zip', ''))
    shutil.make_archive(output_path, "zip", temp_dir)
    shutil.rmtree(temp_dir)
    print(phrase('compression_complete'))

@click.command()
def makeCLI():
    try:
        shutil.rmtree(TEMP_DIR)
    except FileNotFoundError:
        pass
    
    print(phrase('welcome'))
    click.echo(phrase('choose_mode'))
    
    mode = click.getchar()

    match mode:
        case '1':
            print(phrase('basic_mode'))

            minidump_files = glob.glob("C:\\Windows\\Minidump\\*.dmp")

            if not minidump_files:
                print(phrase('no_minidump'))
            else:
                mkdir(TEMP_DIR)
                compress_and_clean(TEMP_DIR, REPORT_ZIP)

        case '2':
            print(phrase('advanced_mode'))
            custom_pause(phrase('pause_message'))
            click.clear()

            mkdir(TEMP_DIR)

            with ThreadPoolExecutor() as executor:
                dxdiag_command = f'dxdiag /t {TEMP_DIR}\\dxdiag_backup.txt'
                wevutil_command = f'wevtutil epl System {TEMP_DIR}\\System.evtx'
                msinfo_command = f'msinfo32 /nfo {TEMP_DIR}\\msinfo32.nfo'
                driverquery_command = (
                    f'driverquery /V > {TEMP_DIR}\\driverV.txt && '
                    f'driverquery /fo csv > {TEMP_DIR}\\driverFo.txt && '
                    f'driverquery /si > {TEMP_DIR}\\driverSi.txt'
                )
                dxdiag_future = executor.submit(run_subprocess, dxdiag_command)
                wevutil_future = executor.submit(run_subprocess, wevutil_command)
                msinfo_future = executor.submit(run_subprocess, msinfo_command)
                driverquery_future = executor.submit(run_subprocess, driverquery_command)

                print(phrase('dxdiag_collecting'))
                dxdiag_future.result()
                print(phrase('dxdiag_success'))

                print(phrase('driver_collecting'))
                driverquery_future.result()
                print(phrase('driver_success'))

                print(phrase('msinfo_collecting'))
                msinfo_future.result()
                print(phrase('msinfo_success'))

            shutil.copyfile("C:\\Windows\\System32\\drivers\\etc\\hosts", f'{TEMP_DIR}\\hosts.txt')
            print(phrase('hosts_success'))

            minidump_files = glob.glob("C:\\Windows\\Minidump\\*.dmp")
            if not minidump_files:
                print(phrase('no_minidump'))
            else:
                with ThreadPoolExecutor() as file_executor:
                    file_executor.map(lambda file: shutil.copyfile(file, f'{TEMP_DIR}\\{path.basename(file)}'), minidump_files)
                print(phrase('minidump_success'))
            compress_and_clean(TEMP_DIR, REPORT_ZIP)

        case _:
            print(phrase('invalid_selection'))
            sys.exit(0)

    custom_pause(phrase('end_message'))

if __name__ == '__main__':
    makeCLI()