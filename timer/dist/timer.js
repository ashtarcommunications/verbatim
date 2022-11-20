import { Timer } from './easytimer/easytimer.js';
// import { Store } from './store.js';

document.addEventListener("DOMContentLoaded", async () => {
    // const store = new Store('.settings.dat');
    // await store.set('some-key', { value: 5 });
    // const val = await store.get('some-key');
    // assert(val, { value: 5 });

    const timer = new Timer();
    let selectedTimer = '#speechtimer';

    const beep = new Audio();
    beep.src = './beep.wav';

    const resetTimers = () => {
        document.querySelector('#activetimer').value = `9:00`;
        document.querySelector('#activetimer').classList = 'speech';
        document.querySelector('#active').classList = 'speech';
        document.querySelector('#afftimer').textContent = `10:00`;
        document.querySelector('#speechtimer').textContent = `9:00`;
        document.querySelector('#negtimer').textContent = `10:00`;
    }
    resetTimers();

    const currentTime = () => {
        const minutes = timer.getTimeValues().minutes;
        let seconds = timer.getTimeValues().seconds;
        seconds = seconds < 10 ? `0${seconds}` : seconds;
        return `${minutes}:${seconds}`;
    };

    const start = () => {
        const value = document.querySelector('#activetimer').value;
        console.log(parseInt(value.split(':')[0]));
        console.log(parseInt(value.split(':')[1]));
        timer.start({
            countdown: true,
            startValues: {
                minutes: parseInt(value.split(':')[0]),
                seconds: parseInt(value.split(':')[1]),
            },
        });
        document.querySelector('#start').style.display = 'none';
        document.querySelector('#pause').style.display = 'block';

        // TODO - better approach to window transparency and turning it off
        // document.querySelector('#smalltimers').style.display = 'none';
        // document.querySelector('#controls').style.display = 'none';
        // document.querySelector('#active').style.backgroundColor = 'transparent';
        // document.querySelector('#activetimer').style.color = 'blue';
    };

    const stop = () => {
        timer.stop();
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

        if (time === '0:30') {
            flashWarning();
        }
    });

    timer.addEventListener('targetAchieved', (e) => {
        flashWarning();
        beep.play();
        stop();
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

    document.querySelector('#showsettings').addEventListener('click', () => {
        document.querySelector('#app').style.display = 'none';
        document.querySelector('#settings').style.display = 'block';
    });
    document.querySelector('#exitsettings').addEventListener('click', () => {
        document.querySelector('#app').style.display = 'block';
        document.querySelector('#settings').style.display = 'none';
    });

    document.querySelector('#affprep').addEventListener('contextmenu', (e) => {
        e.preventDefault();
        return false;
    }, false);
});
