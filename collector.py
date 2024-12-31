import click
from os import path, mkdir
import shutil
import subprocess
import glob
from concurrent.futures import ThreadPoolExecutor
import sys
import locale

TEMP_DIR = path.expanduser("~") + "\\AppData\\Local\\Temp\\SecageReports"

phrases = {
    'welcome': {
        'tr': "Sistem Bilgisi Toplayıcıya hoş geldiniz!",
        'en': "Welcome to the System Information Collector!"
    },
    'choose_mode': {
        'tr': "\nMod seçin: Basit Mod için (1), Gelişmiş Mod için (2) tuşlayın",
        'en': "\nChoose a mode: Press (1) for Simple Mode, (2) for Advanced Mode"
    },
    'invalid_selection': {
        'tr': "\nGeçersiz seçim yapıldı. Program sonlandırılıyor.",
        'en': "\nInvalid selection. The program is terminating."
    },
    'basic_mode': {
        'tr': "\nBasit Mod seçildi: Sadece minidump dosyaları toplanacak.",
        'en': "\nSimple Mode selected: Only minidump files will be collected."
    },
    'no_minidump': {
        'tr': "\nMinidump dosyaları bulunamadı.",
        'en': "\nNo minidump files found."
    },
    'minidump_success': {
        'tr': "\nMinidump dosyaları başarıyla alındı.",
        'en': "\nMinidump files were successfully collected."
    },
    'compressing': {
        'tr': "\nDosyalar sıkıştırılıyor...",
        'en': "\nFiles are being compressed..."
    },
    'compression_complete': {
        'tr': '\nDosyalar sıkıştırıldı ve ZIP dosyası masaüstüne "Raporlar.zip" şeklinde kaydedildi.',
        'en': '\nFiles compressed and the ZIP file was saved to the desktop as "Reports.zip".'
    },
    'advanced_mode': {
        'tr': "\nGelişmiş Mod seçildi.",
        'en': "\nAdvanced Mode selected."
    },
    'pause_message': {
        'tr': "\nLütfen devam etmek için herhangi bir tuşa tıklayın veya 'q' ile çıkın...",
        'en': "\nPlease press any key to continue or 'q' to quit..."
    },
    'hosts_success': {
        'tr': "\nHosts dosyası alındı.",
        'en': "\nHosts file collected."
    },
    'process_error': {
        'tr': "\nİşlem başarısız oldu.",
        'en': "\nProcess failed."
    },
    'goodbye': {
        'tr': "\nÇıkış yapılıyor...",
        'en': "\nExiting..."
    },
    'report_filename': {
        'tr': "Raporlar.zip",
        'en': "Reports.zip"
    },
    'dxdiag_collecting': {
        'tr': "\nDxDiag çıktısı alınıyor...",
        'en': "\nCollecting DxDiag output..."
    },
    'dxdiag_success': {
        'tr': "\nDxDiag çıktısı alındı.",
        'en': "\nDxDiag output collected."
    },
    'driver_collecting': {
        'tr': "\nSürücü listesi alınıyor...",
        'en': "\nCollecting driver list..."
    },
    'driver_success': {
        'tr': "\nSürücü listesi alındı.",
        'en': "\nDriver list collected."
    },
    'msinfo_collecting': {
        'tr': "\nMsinfo32 çıktısı alınıyor...",
        'en': "\nCollecting Msinfo32 output..."
    },
    'msinfo_success': {
        'tr': "\nMsinfo32 çıktısı alındı.",
        'en': "\nMsinfo32 output collected."
    },
    'end_message': {
        'tr': "\nLütfen işlemi sonlandırmak için herhangi bir tuşa tıklayın.",
        'en': "\nPlease press any key to terminate the process."
    }
}

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
    shutil.make_archive(path.expanduser("~") + f"\\Desktop\\{output_name.replace('.zip', '')}", "zip", temp_dir)
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

    if mode == '1':
        print(phrase('basic_mode'))
        
        minidump_files = glob.glob("C:\\Windows\\Minidump\\*.dmp")
        
        if not minidump_files:
            print(phrase('no_minidump'))
        else:
            mkdir(TEMP_DIR)
            compress_and_clean(TEMP_DIR, REPORT_ZIP)

    elif mode == '2':
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

    else:
        print(phrase('invalid_selection'))
        sys.exit(0)

    custom_pause(phrase('end_message'))

if __name__ == '__main__':
    makeCLI()