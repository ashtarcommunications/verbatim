import { Timer } from './easytimer/easytimer.js';
import { Store } from 'tauri-plugin-store-api';
import { appWindow, LogicalSize } from '@tauri-apps/api/window';

document.addEventListener("DOMContentLoaded", async () => {
    const store = new Store('.settings');
    await store.load();
    
    const defaultSettings = {
        alerts: {
            '6:00': false,
            '5:00': true,
            '4:00': false,
            '3:00': false,
            '2:00': false,
            '1:00': false,
            '0:30': true,
            '0:00': true,
        },
        alertTypes: {
            flash: true,
            audio: false,
        },
        speechTimes: {
            constructive: '9:00',
            rebuttal: '6:00',
            cx: '3:00',
            prep: '10:00',
        },
        window: {
            autoshrink: true,
            transparent: true,
            alwaysOnTop: true,
        },
    };
    const savedSettings = await store.get('settings');

    const settings = {...defaultSettings, ...savedSettings};

    await appWindow.setAlwaysOnTop(settings.window.alwaysOnTop);

    document.querySelector('#alert6').checked = settings.alerts['6:00'];
    document.querySelector('#alert5').checked = settings.alerts['5:00'];
    document.querySelector('#alert4').checked = settings.alerts['4:00'];
    document.querySelector('#alert3').checked = settings.alerts['3:00'];
    document.querySelector('#alert2').checked = settings.alerts['2:00'];
    document.querySelector('#alert1').checked = settings.alerts['1:00'];
    document.querySelector('#alert30').checked = settings.alerts['0:30'];
    document.querySelector('#alert0').checked = settings.alerts['0:00'];

    document.querySelector('#warnflash').checked = settings.alertTypes.flash;
    document.querySelector('#warnaudio').checked = settings.alerts.audio;

    document.querySelector('#constructive').value = settings.speechTimes.constructive;
    document.querySelector('#rebuttal').value = settings.speechTimes.rebuttal;
    document.querySelector('#cx').value = settings.speechTimes.cx;
    document.querySelector('#prep').value = settings.speechTimes.prep;

    document.querySelector('#presetconstructive').textContent = settings.speechTimes.constructive;
    document.querySelector('#presetrebuttal').textContent = settings.speechTimes.rebuttal;
    document.querySelector('#presetcx').textContent = settings.speechTimes.cx;
    document.querySelector('#afftimer').textContent = settings.speechTimes.prep;
    document.querySelector('#negtimer').textContent = settings.speechTimes.prep;

    document.querySelector('#autoshrink').checked = settings.window.autoshrink;
    document.querySelector('#transparent').checked = settings.window.transparent;
    document.querySelector('#alwaysontop').checked = settings.window.alwaysOnTop;

    const timer = new Timer();
    let selectedTimer = '#speechtimer';

    const beep = new Audio();
    beep.src = './beep.wav';

    const resetTimers = () => {
        document.querySelector('#activetimer').value = settings.speechTimes.constructive || `9:00`;
        document.querySelector('#activetimer').classList = 'speech';
        document.querySelector('#active').classList = 'speech';
        document.querySelector('#afftimer').textContent = settings.speechTimes.prep || `10:00`;
        document.querySelector('#speechtimer').textContent = settings.speechTimes.constructive || `9:00`;
        document.querySelector('#negtimer').textContent = settings.speechTimes.prep || `10:00`;
    }
    resetTimers();

    const currentTime = () => {
        const minutes = timer.getTimeValues().minutes;
        let seconds = timer.getTimeValues().seconds;
        seconds = seconds < 10 ? `0${seconds}` : seconds;
        return `${minutes}:${seconds}`;
    };

    const start = async () => {
        document.querySelector('#activetimer').disabled = true;
        const value = document.querySelector('#activetimer').value;
        timer.start({
            countdown: true,
            startValues: {
                minutes: parseInt(value.split(':')[0]),
                seconds: parseInt(value.split(':')[1]),
            },
        });

        document.querySelector('#start').style.display = 'none';
        document.querySelector('#pause').style.display = 'block';
        
        if (settings.window.autoshrink) {
            await appWindow.setSize(new LogicalSize(200, 105));
            await appWindow.setDecorations(false);
            document.querySelector('#smalltimers').style.display = 'none';
            document.querySelector('#controls').style.display = 'none';
        }

        if (settings.window.transparent) {
            document.querySelector('#active').classList = 'transparent';
            document.querySelector('#activetimer').classList = 'transparent';
        }
    };

    const stop = async () => {
        timer.stop();
        document.querySelector('#activetimer').disabled = false;
        await appWindow.setSize(new LogicalSize(200, 233));
        await appWindow.setDecorations(true);
        document.querySelector('#smalltimers').style.display = 'flex';
        document.querySelector('#controls').style.display = 'flex';
        switch (selectedTimer) {
            case '#afftimer':
                document.querySelector('#activetimer').classList = 'aff';
                document.querySelector('#active').classList = 'aff';
                break;
            case '#speechtimer':
                document.querySelector('#activetimer').classList = 'speech';
                document.querySelector('#active').classList = 'speech';
                break;
            case '#negtimer':
                document.querySelector('#activetimer').classList = 'neg';
                document.querySelector('#active').classList = 'neg';
                break;
            default:
                break;
        }

        document.querySelector('#pause').style.display = 'none';
        document.querySelector('#start').style.display = 'block';
    }

    const switchTimer = (className) => {
        document.querySelector('#activetimer').classList = className;
        document.querySelector('#active').classList = className;
        switch (className) {
            case 'aff':
                selectedTimer = '#afftimer';
                break;
            case 'speech':
                selectedTimer = '#speechtimer';
                break;
            case 'neg':
                selectedTimer = '#negtimer';
                break;
            default:
                break;
        }
        document.querySelector('#activetimer').value = document.querySelector(selectedTimer).textContent;
    }

    const flashWarning = () => {
        const currentClass = document.querySelector('#activetimer').classList.toString();
        let repeat = 0;
        const clear = setInterval(() => {
            if (document.querySelector('#activetimer').classList.toString() === currentClass) {
                document.querySelector('#activetimer').classList = 'warn';
                document.querySelector('#active').classList = 'warn';
            } else {
                document.querySelector('#activetimer').classList = currentClass;
                document.querySelector('#active').classList = currentClass;
            }
            if (repeat++ === 5) {
                clearInterval(clear);
                document.querySelector('#activetimer').classList = currentClass;
                document.querySelector('#active').classList = currentClass;
            }
        }, 200);
    }

    timer.addEventListener('secondsUpdated', () => {
        const time = currentTime();
        document.querySelector('#activetimer').value = time;
        document.querySelector(selectedTimer).textContent = time;

        if (Object.keys(settings.alerts).includes(time) && settings.alerts[time]) {
            if (window.settings.alertTypes.flash) {
                flashWarning();
            }
            if (window.settings.alertTypes.beep) {
                beep.play();
            }
        }
    });

    timer.addEventListener('targetAchieved', async () => {
        await stop();
        if (settings.alerts['0:00']) {
            if (window.settings.alertTypes.flash) {
                flashWarning();
            }
            if (window.settings.alertTypes.beep) {
                beep.play();
            }
        }
    });

    document.querySelector('#start').addEventListener('click', () => {
        validate();
        start();
    });

    document.querySelector('#pause').addEventListener('click', () => {
        stop();
    });

    document.querySelector('#active').addEventListener('click', () => {
        timer.isRunning ? stop() : start();
    });

    document.querySelector('#affprep').addEventListener('click', () => {
        stop();
        switchTimer('aff');
    });

    document.querySelector('#speech').addEventListener('click', () => {
        stop();
        switchTimer('speech');

    });

    document.querySelector('#negprep').addEventListener('click', () => {
        stop();
        switchTimer('neg');
    });

    document.querySelector('#reset').addEventListener('click', () => {
        document.querySelector('#confirmreset').style.display = 'block';
        document.querySelector('#reset svg').style.display = 'none';
        setTimeout(() => {
            document.querySelector('#confirmreset').style.display = 'none';
            document.querySelector('#reset svg').style.display = 'block';
        }, 3000);
    });

    document.querySelector('#confirmreset').addEventListener('click', (e) => {
        e.stopPropagation(); // Prevents the reset div click handler firing again
        resetTimers();
        document.querySelector('#confirmreset').style.display = 'none';
        document.querySelector('#reset svg').style.display = 'block';
    });

    const validate = () => {
        const value = document.querySelector('#activetimer').value.replace(/\D/g, '');
        const seconds = value.length < 3 ? parseInt(value) : parseInt(value.slice(-2));
        const minutes = value.length > 2
            ? value.length === 3
                ? parseInt(value[0])
                : parseInt(value.substring(0, 2))
            : 0;
        let timeString = '0:00';

        switch (value.length) {
            case 0:
                break;
            case 1:
                timeString = `0:${value}`;
            case 2:
                if (seconds <= 59) {
                    timeString = `0:${value}`;
                } else {
                    timeString = `1:${seconds - 60 < 10 ? `0${seconds - 60}` : seconds - 60}`;
                }
                break;
            case 3:
                timeString = `${value[0]}:${seconds > 59 ? 59 : seconds < 10 ? `0${seconds}` : seconds}`;
                break;
            case 4:
                timeString = `${minutes > 59 ? 59 : minutes}:${seconds > 59 ? 59 : seconds < 10 ? `0${seconds}` : seconds}`;
                break;
            default:
                break;
        }
        document.querySelector('#activetimer').value = timeString;
    };

    document.querySelector('#activetimer').addEventListener('change', () => {
        validate();
    });

    Array.from(document.getElementsByClassName('presettime')).forEach(element => {
        element.addEventListener('click', (e) => {
            stop();
            document.querySelector('#activetimer').value = e.target.textContent;
            document.querySelector('#speechtimer').textContent = e.target.textContent;
            switchTimer('speech');
        });
    });

    document.querySelector('#showsettings').addEventListener('click', async () => {
        // I do not know why this is necessary, but apparently it is
        document.body.style.height = 'auto';
        
        // Swap to settings view
        document.querySelector('#app').style.display = 'none';
        document.querySelector('#settings').style.display = 'block';
    });

    document.querySelector('#presettimes').addEventListener('change', () => {
        const value = document.querySelector('#presettimes').value;
        switch (value) {
            case '':
                break;
            case 'ccx':
                document.querySelector('#constructive').value = '9:00';
                document.querySelector('#rebuttal').value = '6:00';
                document.querySelector('#cx').value = '3:00';
                document.querySelector('#prep').value = '10:00';
                break;
            case 'hscx':
                document.querySelector('#constructive').value = '8:00';
                document.querySelector('#rebuttal').value = '5:00';
                document.querySelector('#cx').value = '3:00';
                document.querySelector('#prep').value = '8:00';
                break;
            case 'hsld':
                document.querySelector('#constructive').value = '6:00';
                document.querySelector('#rebuttal').value = '4:00';
                document.querySelector('#cx').value = '3:00';
                document.querySelector('#prep').value = '4:00';
                break;
            case 'hspf':
                document.querySelector('#constructive').value = '4:00';
                document.querySelector('#rebuttal').value = '3:00';
                document.querySelector('#cx').value = '3:00';
                document.querySelector('#prep').value = '3:00';
                break;
            default:
                break;
        }
    });

    document.querySelector('#savesettings').addEventListener('click', async () => {
        settings.alerts['6:00'] = document.querySelector('#alert6').checked;
        settings.alerts['5:00'] = document.querySelector('#alert5').checked;
        settings.alerts['4:00'] = document.querySelector('#alert4').checked;
        settings.alerts['3:00'] = document.querySelector('#alert3').checked;
        settings.alerts['2:00'] = document.querySelector('#alert2').checked;
        settings.alerts['1:00'] = document.querySelector('#alert1').checked;
        settings.alerts['0:30'] = document.querySelector('#alert30').checked;
        settings.alerts['0:00'] = document.querySelector('#alert0').checked;

        settings.alertTypes.flash = document.querySelector('#warnflash').checked;
        settings.alertTypes.audio = document.querySelector('#warnaudio').checked;

        // TODO - normalize values;
        settings.speechTimes.constructive = document.querySelector('#constructive').value;
        settings.speechTimes.rebuttal = document.querySelector('#rebuttal').value;
        settings.speechTimes.cx = document.querySelector('#cx').value;
        settings.speechTimes.prep = document.querySelector('#prep').value;

        settings.window.autoshrink = document.querySelector('#autoshrink').checked;
        settings.window.transparent = document.querySelector('#transparent').checked;
        settings.window.alwaysOnTop = document.querySelector('#alwaysontop').checked;

        store.save();

        // I, again, do not know why this is necessary, but it works
        document.body.style.height = '100%';

        // Set preset speech times to match new settings
        document.querySelector('#presetconstructive').textContent = settings.speechTimes.constructive;
        document.querySelector('#presetrebuttal').textContent = settings.speechTimes.rebuttal;
        document.querySelector('#presetcx').textContent = settings.speechTimes.cx;
        
        // Switch back to main app view
        document.querySelector('#app').style.display = 'flex';
        document.querySelector('#settings').style.display = 'none';

        await appWindow.setAlwaysOnTop(settings.window.alwaysOnTop);
    });

    document.querySelector('#affprep').addEventListener('contextmenu', (e) => {
        e.preventDefault();
        return false;
    }, false);
});
