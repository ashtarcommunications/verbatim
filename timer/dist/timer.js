/* eslint no-unused-vars: "off", no-undef: "off", no-new: "off" */
// import * as  Timer from './easytimer.js';
import { Timer } from './easytimer/easytimer.js';

export const countdownString = timer => {
    const minutes = timer.getTimeValues().minutes < 10 ? `0${timer.getTimeValues().minutes}` : timer.getTimeValues().minutes;
    const seconds = timer.getTimeValues().seconds < 10 ? `0${timer.getTimeValues().seconds}` : timer.getTimeValues().seconds;
    return `${minutes}:${seconds}`;
};

export const setupTimer = () => {
    const timer = new Timer();
    const getTimerForm = document.getElementById('timerForm');
    const getTimerDisplay = document.getElementById('mainTimerDisplay');
    const getTimerSubmit = document.getElementById('timerSubmit');
    const getMinutes = () => parseInt(document.getElementById('minuteSelect').value);
    const getSeconds = () => parseInt(document.getElementById('secondSelect').value);

    timer.addEventListener('secondsUpdated', () => (
        getTimerDisplay.textContent = countdownString(timer)
    ));

    const startCountdown = () => {
        timer.start({
            countdown: true,
            startValues: {
                seconds: getSeconds(),
                minutes: getMinutes(),
            },
        });
    };

    getTimerForm.addEventListener('submit', event => {
        event.preventDefault();
        const submitTextValue = getTimerSubmit.textContent;
        switch (submitTextValue) {
            case 'start':
                startCountdown();
                getTimerSubmit.textContent = 'pause';
                break;
            case 'pause':
                timer.pause();
                getTimerSubmit.textContent = 'resume';
                break;
            case 'resume':
                timer.start();
                getTimerSubmit.textContent = 'pause';
                break;
            default:
                timer.pause();
                break;
        }
    });

    getTimerForm.addEventListener('reset', event => {
        event.preventDefault();
        timer.reset();
        timer.stop();
        getTimerSubmit.textContent = 'start';
        getTimerDisplay.textContent = '00:00';
    });
};

export const timeOperations = (timer, classList) => {
    timer.pause();
    if (classList.contains('plus-minute')) {
        const newMin = timer.getTimeValues().minutes + 1;
        if (newMin <= 59) timer.getTimeValues().minutes = newMin;
    }
    if (classList.contains('minus-minute')) {
        const newMin = timer.getTimeValues().minutes - 1;
        if (newMin >= 0) timer.getTimeValues().minutes = newMin;
    }
    if (classList.contains('plus-second')) {
        const newSec = timer.getTimeValues().seconds + 1;
        if (newSec <= 59) timer.getTimeValues().seconds = newSec;
    }
    if (classList.contains('minus-second')) {
        const newSec = timer.getTimeValues().seconds - 1;
        if (newSec >= 0) timer.getTimeValues().seconds = newSec;
    }
    return new Timer({
        countdown: true,
        startValues: [0, timer.getTimeValues().seconds, timer.getTimeValues().minutes, 0, 0],
    });
};

export const setupPrepTimers = () => {
    const handleTimerState = (timer, form) => {
        const submitButton = form.querySelector('button[type="submit"]');
        switch (submitButton.textContent) {
            case 'start':
                timer.start();
                submitButton.textContent = 'pause';
                break;
            case 'pause':
                timer.pause();
                submitButton.textContent = 'resume';
                break;
            case 'resume':
                timer.start();
                submitButton.textContent = 'pause';
                break;
            default:
                timer.pause();
                break;
        }
    };

    const setPrepTimerDisplay = form => form.querySelector('.timer-display');

    let firstPrepTimer = new Timer({ countdown: true, startValues: [0, 0, 3, 0, 0] });
    const firstPrepTimerForm = document.querySelector('#firstPrepTimerForm');
    const setFirstPrepTimerDisplay = () => (
        setPrepTimerDisplay(firstPrepTimerForm).textContent = countdownString(firstPrepTimer)
    );
    setFirstPrepTimerDisplay();

    firstPrepTimerForm.addEventListener('submit', event => {
        event.preventDefault();
        handleTimerState(firstPrepTimer, event.target);
    });
    firstPrepTimerForm.addEventListener('reset', event => {
        event.preventDefault();
        firstPrepTimer = new Timer({ countdown: true, startValues: [0, 0, 3, 0, 0] });
        setFirstPrepTimerDisplay();
        firstPrepTimerForm.querySelector('button[type="submit"]').textContent = 'start';
    });

    Array.from(
        firstPrepTimerForm.querySelectorAll('.minus-second, .minus-minute, .plus-second, .plus-minute')
    )
    .forEach(element => element.addEventListener('click', event => {
        event.preventDefault();
        event.target.form.querySelector('button[type="submit"]').textContent = 'resume';

        firstPrepTimer.removeEventListener('secondsUpdated', setFirstPrepTimerDisplay);
        firstPrepTimer = timeOperations(firstPrepTimer, event.target.classList);
        setFirstPrepTimerDisplay();
        firstPrepTimer.addEventListener('secondsUpdated', setFirstPrepTimerDisplay);
    })
    );

    firstPrepTimer.addEventListener('secondsUpdated', setFirstPrepTimerDisplay);

    let secondPrepTimer = new Timer({ countdown: true, startValues: [0, 0, 3, 0, 0] });
    const secondPrepTimerForm = document.querySelector('#secondPrepTimerForm');
    const setSecondPrepTimerDisplay = () => (
        // eslint-disable-next-line max-len
        setPrepTimerDisplay(secondPrepTimerForm).textContent = countdownString(secondPrepTimer)
    );
    setSecondPrepTimerDisplay();

    secondPrepTimerForm.addEventListener('submit', event => {
        event.preventDefault();
        handleTimerState(secondPrepTimer, event.target);
    });
    secondPrepTimerForm.addEventListener('reset', event => {
        event.preventDefault();
        secondPrepTimer = new Timer({ countdown: true, startValues: [0, 0, 3, 0, 0] });
        setSecondPrepTimerDisplay();
        secondPrepTimerForm.querySelector('button[type="submit"]').textContent = 'start';
    });

    Array.from(
        secondPrepTimerForm.querySelectorAll('.minus-second, .minus-minute, .plus-second, .plus-minute')
    )
    .forEach(element => element.addEventListener('click', event => {
        event.preventDefault();
        event.target.form.querySelector('button[type="submit"]').textContent = 'resume';

        secondPrepTimer.removeEventListener('secondsUpdated', setSecondPrepTimerDisplay);
        secondPrepTimer = timeOperations(secondPrepTimer, event.target.classList);
        setSecondPrepTimerDisplay();
        secondPrepTimer.addEventListener('secondsUpdated', setSecondPrepTimerDisplay);
    })
    );

    secondPrepTimer.addEventListener('secondsUpdated', setSecondPrepTimerDisplay);

    const setPrepTimersDisplay = () => {
        const prepTimers = document.getElementById('prepTimers');
        /* eslint-disable no-unused-expressions */
        document.getElementById('prepTimersCheckbox').checked ?
            prepTimers.classList.remove('no-debate-timers')
        :
            prepTimers.classList.add('no-debate-timers');
        /* eslint-enable no-unused-expressions */
    };
    setPrepTimersDisplay();

    document.getElementById('prepTimersCheckbox').addEventListener('change', setPrepTimersDisplay);
};
