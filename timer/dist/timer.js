/* eslint-disable no-nested-ternary, no-plusplus, import/no-unresolved */
/* global document, Audio */
import { Store } from 'tauri-plugin-store-api';
import { appWindow, PhysicalSize } from '@tauri-apps/api/window';
import { Timer } from './easytimer/easytimer.js';

// Set up a persistent store for settings, file location is in OS-specific config directory
const store = new Store('.settings');
let settings = {};

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
    sideNames: 'aff',
    window: {
        defaultWidth: 200,
        defaultHeight: 200,
        currentWidth: 200,
        currentHeight: 200,
        autoshrink: true,
        transparent: true,
        alwaysOnTop: true,
        transparencyColor: 'white',
    },
    currentTimes: {
        speech: '9:00',
        aff: '10:00',
        neg: '10:00',
    },
};

// Initialize timer - we only need one, because timer state for each timer
// is naturally stored in the DOM by updating the value on page
const timer = new Timer();
let selectedTimer = '#speechtimer';

const beep = new Audio();
beep.src = './beep.wav';

appWindow.onCloseRequested(async () => {
    // Persist the current timer state in case of accidentally closing app
    settings.currentTimes.speech = document.querySelector('#speechtimer').textContent;
    settings.currentTimes.aff = document.querySelector('#afftimer').textContent;
    settings.currentTimes.neg = document.querySelector('#negtimer').textContent;
    await store.set('settings', settings);
    await store.save();
});

appWindow.onResized(async ({ payload: size }) => {
    // Save resized window so we can restore on open
    settings.window.defaultWidth = size.width;
    settings.window.defaultHeight = size.height;
});

const setSideNames = () => {
    switch (settings.sideNames) {
        case 'aff':
            document.querySelector('#afflabel').textContent = 'Aff Prep';
            document.querySelector('#neglabel').textContent = 'Neg Prep';
            break;
        case 'pro':
            document.querySelector('#afflabel').textContent = 'Pro Prep';
            document.querySelector('#neglabel').textContent = 'Con Prep';
            break;
        case 'gov':
            document.querySelector('#afflabel').textContent = 'Gov Prep';
            document.querySelector('#neglabel').textContent = 'Opp Prep';
            break;
        default:
            break;
    }
};

const setupDefaultSettings = async () => {
    try {
        await store.load();
    } catch (err) {
        // An error is thrown if settings file doesn't exist, so save to generate
        // file if the load call fails
        await store.save();
    }

    // Merge saved settings
    const savedSettings = await store.get('settings');
    settings = { ...defaultSettings, ...savedSettings };

    // Restore window to saved size
    await appWindow.setAlwaysOnTop(settings.window.alwaysOnTop);
    await appWindow.setSize(new PhysicalSize(
        settings.window.defaultWidth || 200,
        settings.window.defaultHeight || 200
    ));

    // Make the settings page match the saved settings
    document.querySelector('#alert6').checked = settings.alerts['6:00'];
    document.querySelector('#alert5').checked = settings.alerts['5:00'];
    document.querySelector('#alert4').checked = settings.alerts['4:00'];
    document.querySelector('#alert3').checked = settings.alerts['3:00'];
    document.querySelector('#alert2').checked = settings.alerts['2:00'];
    document.querySelector('#alert1').checked = settings.alerts['1:00'];
    document.querySelector('#alert30').checked = settings.alerts['0:30'];
    document.querySelector('#alert0').checked = settings.alerts['0:00'];

    document.querySelector('#warnflash').checked = settings.alertTypes.flash;
    document.querySelector('#warnaudio').checked = settings.alertTypes.audio;

    document.querySelector('#constructive').value = settings.speechTimes.constructive;
    document.querySelector('#rebuttal').value = settings.speechTimes.rebuttal;
    document.querySelector('#cx').value = settings.speechTimes.cx;
    document.querySelector('#prep').value = settings.speechTimes.prep;

    document.querySelector('#sidenames').value = settings.sideNames;

    document.querySelector('#autoshrink').checked = settings.window.autoshrink;
    document.querySelector('#transparent').checked = settings.window.transparent;
    document.querySelector('#alwaysontop').checked = settings.window.alwaysOnTop;
    document.querySelector('#transparencycolor').value = settings.window.transparencyColor;

    // Make the timer UI match the settings
    document.querySelector('#presetconstructive').textContent = settings.speechTimes.constructive;
    document.querySelector('#presetrebuttal').textContent = settings.speechTimes.rebuttal;
    document.querySelector('#presetcx').textContent = settings.speechTimes.cx;
    setSideNames();

    // Restore saved timer state
    document.querySelector('#activetimer').value = settings.currentTimes.speech || settings.speechTimes.speech;
    document.querySelector('#speechtimer').textContent = settings.currentTimes.speech || settings.speechTimes.speech;
    document.querySelector('#afftimer').textContent = settings.currentTimes.aff || settings.speechTimes.prep;
    document.querySelector('#negtimer').textContent = settings.currentTimes.neg || settings.speechTimes.prep;
};

const resetTimers = async () => {
    // Reset the active timer and make speech timer active
    document.querySelector('#activetimer').value = settings.speechTimes.constructive || '9:00';
    document.querySelector('#activetimer').classList = 'speech';
    document.querySelector('#active').classList = 'speech';
    document.querySelector('#speechtimer').textContent = settings.speechTimes.constructive || '9:00';
    selectedTimer = '#speechtimer';

    // Reset prep timers
    document.querySelector('#afftimer').textContent = settings.speechTimes.prep || '10:00';
    document.querySelector('#negtimer').textContent = settings.speechTimes.prep || '10:00';

    // Clear any saved timer state
    settings.currentTimes.speech = settings.speechTimes.constructive || '9:00';
    settings.currentTimes.aff = settings.speechTimes.prep || '10:00';
    settings.currentTimes.neg = settings.speechTimes.prep || '10:00';
    store.set('settings', settings);
    store.save();
};

const currentTime = () => {
    // Convert the internal timer state to a display string
    const minutes = timer.getTimeValues().minutes;
    let seconds = timer.getTimeValues().seconds;
    seconds = seconds < 10 ? `0${seconds}` : seconds;
    return `${minutes}:${seconds}`;
};

const start = async () => {
    // Prevent time entry while running
    document.querySelector('#activetimer').disabled = true;

    // Run the actual timer
    const value = document.querySelector('#activetimer').value;
    timer.start({
        countdown: true,
        startValues: {
            minutes: parseInt(value.split(':')[0]),
            seconds: parseInt(value.split(':')[1]),
        },
    });

    // Swap start/pause buttons
    document.querySelector('#start').style.display = 'none';
    document.querySelector('#pause').style.display = 'block';

    if (settings.window.autoshrink) {
        // Get the current window size to restore it on stop
        const currentSize = await appWindow.innerSize();
        settings.window.currentWidth = currentSize.width;
        settings.window.currentHeight = currentSize.height;
        await appWindow.setSize(new PhysicalSize(
            currentSize.width || 200,
            document.querySelector('#active').offsetHeight || 105
        ));

        // Hide rest of window except the timer
        await appWindow.setDecorations(false);
        document.querySelector('#smalltimers').style.display = 'none';
        document.querySelector('#controls').style.display = 'none';
    }

    if (settings.window.transparent) {
        document.querySelector('#active').classList = `transparent ${settings.window.transparencyColor}`;
        document.querySelector('#activetimer').classList = `transparent ${settings.window.transparencyColor}`;
    }
};

const stop = async () => {
    timer.stop();
    document.querySelector('#activetimer').disabled = false; // Reenable time entry

    // Restore the window to pre-shrunk size
    if (settings.window.autoshrink) {
        await appWindow.setDecorations(true);
        await appWindow.setSize(new PhysicalSize(
            settings.window.currentWidth || 200,
            settings.window.currentHeight || 200
        ));       
        
        document.querySelector('#smalltimers').style.display = 'flex';
        document.querySelector('#controls').style.display = 'flex';
    }

    // Remove transparency
    if (settings.window.transparent) {
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
    }

    // Swap pause/start buttons
    document.querySelector('#pause').style.display = 'none';
    document.querySelector('#start').style.display = 'block';
};

const switchTimer = (className) => {
    // Change the main timer UI to match color of selected timer
    document.querySelector('#activetimer').classList = className;
    document.querySelector('#active').classList = className;

    // Switch the active timer
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
};

const flashAlert = async () => {
    const currentClass = document.querySelector('#activetimer').classList.toString();
    let repeat = 0;
    const clear = setInterval(() => {
        if (document.querySelector('#activetimer').classList.toString() === currentClass) {
            document.querySelector('#activetimer').classList = 'warn';
        } else {
            document.querySelector('#activetimer').classList = currentClass;
        }
        // eslint-disable-next-line no-plusplus
        if (repeat++ > 4) {
            clearInterval(clear);
            document.querySelector('#activetimer').classList = currentClass;
        }
    }, 200);
};

const cleanTimeInput = (inputString) => {
    // Strip non-numeric characters
    const value = inputString.replace(/\D/g, '');

    // Interpret 1 or 2 characters as seconds, if 3 or 4 characters take the last 2
    const seconds = value.length < 3 ? parseInt(value) : parseInt(value.slice(-2));

    // If 3 characters, the first is the minute, if more than 3, take the first 2
    // This will drop the middle character if 5 digits are entered
    // we'll assume they meant a colon
    const minutes = value.length > 2
        ? value.length === 3
            ? parseInt(value[0])
            : parseInt(value.substring(0, 2))
        : 0;

    // Default to zero if we can't parse
    let timeString = '0:00';

    switch (value.length) {
        case 0:
            break;
        case 1:
            // Only seconds entered
            timeString = `0:${value}`;
            break;
        case 2:
            // If less than a minute, use seconds, otherwise convert seconds > 60 to minutes
            // This allows shortcuts like entering '90' for 1:30
            if (seconds <= 59) {
                timeString = `0:${value}`;
            } else {
                timeString = `1:${seconds - 60 < 10 ? `0${seconds - 60}` : seconds - 60}`;
            }
            break;
        case 3:
            // x:xx entered, but override invalid values > 59
            timeString = `${value[0]}:${seconds > 59 ? 59 : seconds < 10 ? `0${seconds}` : seconds}`;
            break;
        case 4:
            // xx:xx entered, but override invalid values > 59
            timeString = `${minutes > 59 ? 59 : minutes}:${seconds > 59 ? 59 : seconds < 10 ? `0${seconds}` : seconds}`;
            break;
        case 5:
            // xxxxx entered, treat it the same as 4 characters since we already dropped the middle
            timeString = `${minutes > 59 ? 59 : minutes}:${seconds > 59 ? 59 : seconds < 10 ? `0${seconds}` : seconds}`;
            break;
        default:
            break;
    }

    return timeString;
};

const setTimer = (time) => {
    // Set both the active timer and small timer for the provided time
    document.querySelector('#activetimer').value = time;
    switch (selectedTimer) {
        case '#afftimer':
            document.querySelector('#afftimer').textContent = time;
            break;
        case '#speechtimer':
            document.querySelector('#speechtimer').textContent = time;
            break;
        case '#negtimer':
            document.querySelector('#negtimer').textContent = time;
            break;
        default:
            break;
    }
};

document.addEventListener('DOMContentLoaded', async () => {
    // App initialization
    await setupDefaultSettings();

    // Main timer loop that runs every second
    timer.addEventListener('secondsUpdated', () => {
        // Get the time from the running timer
        const time = currentTime();

        // Update the main timer and small timer to match the timer state
        document.querySelector('#activetimer').value = time;
        document.querySelector(selectedTimer).textContent = time;

        // Run any configured alerts on matching times, except 0:00 which we'll catch separately
        if (settings.alerts[time] && time !== '0:00') {
            if (settings.alertTypes.flash) {
                flashAlert();
            }
            if (settings.alertTypes.beep) {
                beep.play();
            }
        }
    });

    // Runs at 0:00
    timer.addEventListener('targetAchieved', async () => {
        // Run the final alerts here to avoid a race condition with stopping the timer
        if (settings.alerts['0:00']) {
            if (settings.alertTypes.flash) {
                await flashAlert();
            }
            if (settings.alertTypes.beep) {
                beep.play();
            }
        }
        await stop();
    });

    document.querySelector('#start').addEventListener('click', () => {
        // Validate any user input to ensure a valid time, then start
        const newTime = cleanTimeInput(document.querySelector('#activetimer').value);
        setTimer(newTime);
        start();
    });

    document.querySelector('#pause').addEventListener('click', () => {
        stop();
    });

    document.querySelector('#active').addEventListener('click', () => {
        // Have to allow clicking on the timer itself to stop if pause button is hidden
        if (timer.isRunning()) { stop(); }
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
        // Display a confirmation button for 3 seconds or revert
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

    document.querySelector('#activetimer').addEventListener('change', () => {
        // User input in the main timer, make sure it's a valid time then update timer
        const newTime = cleanTimeInput(document.querySelector('#activetimer').value);
        setTimer(newTime);
    });

    Array.from(document.getElementsByClassName('presettime')).forEach(element => {
        // Switch to the speech timer with the selected preset time
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
            case 'ccx': // College CX
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
        // Update the global settings with the new values
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

        settings.speechTimes.constructive = cleanTimeInput(document.querySelector('#constructive').value);
        settings.speechTimes.rebuttal = cleanTimeInput(document.querySelector('#rebuttal').value);
        settings.speechTimes.cx = cleanTimeInput(document.querySelector('#cx').value);
        settings.speechTimes.prep = cleanTimeInput(document.querySelector('#prep').value);

        settings.sideNames = document.querySelector('#sidenames').value;

        settings.window.autoshrink = document.querySelector('#autoshrink').checked;
        settings.window.transparent = document.querySelector('#transparent').checked;
        settings.window.alwaysOnTop = document.querySelector('#alwaysontop').checked;
        settings.window.transparencyColor = document.querySelector('#transparencycolor').value;

        await store.set('settings', settings);
        await store.save();

        // Update the inputs on the settings page with the cleaned values
        document.querySelector('#constructive').value = settings.speechTimes.constructive;
        document.querySelector('#rebuttal').value = settings.speechTimes.rebuttal;
        document.querySelector('#cx').value = settings.speechTimes.cx;
        document.querySelector('#prep').value = settings.speechTimes.prep;

        // I, again, do not know why this is necessary, but it works
        document.body.style.height = '100%';

        // Update timer UI to match new preset settings
        document.querySelector('#presetconstructive').textContent = settings.speechTimes.constructive;
        document.querySelector('#presetrebuttal').textContent = settings.speechTimes.rebuttal;
        document.querySelector('#presetcx').textContent = settings.speechTimes.cx;

        // Update text on prep timers
        setSideNames();

        // Switch back to main app view
        document.querySelector('#app').style.display = 'flex';
        document.querySelector('#settings').style.display = 'none';

        await appWindow.setAlwaysOnTop(settings.window.alwaysOnTop);
    });

    // Right-click menu for the aff prep timer
    document.querySelector('#affprep').addEventListener('contextmenu', (e) => {
        e.preventDefault();
        if (timer.isRunning() && !timer.isPaused()) { return false; }
        document.querySelector('#affprepdisplay').style.display = 'none';
        document.querySelector('#affprepmenu').style.display = 'flex';
        setTimeout(() => {
            document.querySelector('#affprepdisplay').style.display = 'block';
            document.querySelector('#affprepmenu').style.display = 'none';
        }, 3000);
        return false;
    }, false);

    // Right-click menu for the neg prep timer
    document.querySelector('#negprep').addEventListener('contextmenu', (e) => {
        e.preventDefault();
        if (timer.isRunning() && !timer.isPaused()) { return false; }
        document.querySelector('#negprepdisplay').style.display = 'none';
        document.querySelector('#negprepmenu').style.display = 'flex';
        setTimeout(() => {
            document.querySelector('#negprepdisplay').style.display = 'block';
            document.querySelector('#negprepmenu').style.display = 'none';
        }, 3000);
        return false;
    }, false);

    document.querySelector('#affprepreset').addEventListener('click', (e) => {
        e.stopPropagation(); // Prevents switching timers

        document.querySelector('#afftimer').textContent = settings.speechTimes.prep || '10:00';
        if (selectedTimer === '#afftimer') {
            document.querySelector('#activetimer').value = settings.speechTimes.prep || '10:00';
        }
        document.querySelector('#affprepdisplay').style.display = 'block';
        document.querySelector('#affprepmenu').style.display = 'none';
    });

    document.querySelector('#negprepreset').addEventListener('click', (e) => {
        e.stopPropagation(); // Prevents switching timers

        document.querySelector('#negtimer').textContent = settings.speechTimes.prep || '10:00';
        if (selectedTimer === '#negtimer') {
            document.querySelector('#activetimer').value = settings.speechTimes.prep || '10:00';
        }
        document.querySelector('#negprepdisplay').style.display = 'block';
        document.querySelector('#negprepmenu').style.display = 'none';
    });
});
