// ==UserScript==
// @name         Microsites_OOTO
// @namespace    https://amazon.com/
// @version      0.2
// @description  Schedule Quick connects and send OOTO from the phonetool page - Microsites
// @author       @mofila
// @match        https://phonetool.amazon.com/users/*
// @match        https://connect.amazon.com/users/*
// @match        https://outlook.office.com/*
// @updateURL    https://raw.githubusercontent.com/Mofi-l/Microsites_OOTO/main/Microsites_OOTO_meta.js
// @downloadURL  https://raw.githubusercontent.com/Mofi-l/Microsites_OOTO/main/Microsites_OOTO_user.js
// @grant        GM.xmlHttpRequest
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        unsafeWindow
// ==/UserScript==

/* globals $, moment */

(function() {
    'use strict';

    const MEETING_TYPES = {
        OOTO: 'ooto',
        QUICK: 'quick'
    };

    function getFormattedDateRange(start, end) {
        return `(${moment(start).format('YYYY-MM-DD')} - ${moment(end).format('YYYY-MM-DD')})`;
    }

    const buildDataProvider = (name, endpoint) => ({
        name: name,
        endpoint: endpoint,
        onajaxerror (reject) {
            return _ => {
                console.error(_);
                reject(_);
            };
        },

        fetchEmail (username, isRetry) {
            return new Promise((resolve, reject) => {
                GM.xmlHttpRequest({
                    method: 'POST',
                    url: `${endpoint}/owa/service.svc`,
                    headers: {
                        action: 'FindPeople',
                        'content-type': 'application/json; charset=UTF-8',
                        'x-owa-actionname': 'OwaOptionPage',
                        'x-owa-canary': this.getToken(),
                        ...(this.getJwt() && { 'authorization': this.getJwt() })
                    },
                    data: JSON.stringify({
                        Header: {
                            RequestServerVersion: 'Exchange2013'
                        },
                        Body: {
                            IndexedPageItemView: {
                                __type: 'IndexedPageView:#Exchange',
                                BasePoint: 'Beginning'
                            },
                            QueryString: username + '@'
                        }
                    }),
                    onerror: this.onajaxerror(reject),
                    onload: _ => {
                        this.updateToken(_);
                        if (_.status == 200) {
                            const responseBody = JSON.parse(_.response).Body;
                            var amazonAddress = responseBody.ResultSet
                            .filter(_ => Object.keys(_).includes('PersonaTypeString'))
                            .filter(_ => _.PersonaTypeString == 'Person')
                            .filter(_ => Object.keys(_).includes('EmailAddress'))
                            .map(_ => _.EmailAddress.EmailAddress)
                            .find(_ => _.startsWith(username + '@amazon'));
                            amazonAddress = amazonAddress || username + '@amazon.com';
                            resolve(amazonAddress);
                        } else {
                            this.onajaxerror(reject)(_);
                        }
                    }
                });
            }).catch(e => {
                if (isRetry) {
                    throw e;
                }
                return this.fetchEmail(username, true);
            });
        },

        getAvailability (mailboxes, start, end, currentUserEmail, isSelfView) {
            return new Promise((resolve, reject) => {
                GM.xmlHttpRequest({
                    method: 'POST',
                    url: `${endpoint}/owa/service.svc`,
                    headers: {
                        action: 'GetUserAvailabilityInternal',
                        'content-type': 'application/json; charset=UTF-8',
                        'x-owa-actionname': 'GetUserAvailabilityInternal_FetchWorkingHours',
                        'x-owa-canary': this.getToken(),
                        ...(this.getJwt() && { 'authorization': this.getJwt() })
                    },
                    data: JSON.stringify({
                        request: {
                            Header: {
                                RequestServerVersion: 'Exchange2013',
                                TimeZoneContext: {
                                    TimeZoneDefinition: { Id: 'UTC' }
                                }
                            },
                            Body: {
                                MailboxDataArray: mailboxes.map(_ => ({ Email: { Address: _ }})),
                                FreeBusyViewOptions: {
                                    RequestedView: 'Detailed',
                                    TimeWindow: { StartTime: start.toISOString(), EndTime: end.toISOString() }
                                }
                            }
                        }
                    }),
                    onerror: this.onajaxerror(reject),
                    onload: _ => {
                        this.updateToken(_);
                        if (_.status == 200) {
                            const response = JSON.parse(_.response);
                            const events = [];
                            response.Body.Responses.map(_ => _.CalendarView.Items.filter(_ => _.FreeBusyType != 'Free')).forEach(group => {
                                for (let i = 0; i < group.length; i++) {
                                    if (!isSelfView) {
                                        if (group[i + 1] && group[i].End >= group[i + 1].Start) {
                                            if (group[i].FreeBusyType == group[i + 1].FreeBusyType) {
                                                group[i + 1].Start = group[i].Start;
                                                if (group[i].End > group[i + 1].End) {
                                                    group[i + 1].End = group[i].End;
                                                }
                                                continue;
                                            }
                                        }
                                    }
                                    events.push(group[i]);
                                }
                            });
                            resolve(events.map(_ => {
                                let status = _.FreeBusyType.toLowerCase();
                                let subject = _.Subject ? _.Subject : "No Title";
                                if (status == 'oof') {
                                    status = 'out-of-office';
                                }
                                if (shouldShowHireEvents() && (subject.includes("interview") || subject.includes("Debrief") || subject.includes("Pre-brief"))) {
                                    status = "hire";
                                }
                                return {
                                    className: 'calendar-' + (_.ParentFolderId.Id == currentUserEmail ? 'my-' : '') + status,
                                    title: isSelfView ? subject : status,
                                    start: moment.utc(_.Start),
                                    end: moment.utc(_.End)
                                };
                            }));
                        } else {
                            this.onajaxerror(reject)(_);
                        }
                    }
                });
            });
        },

        createMeeting({ meetingType, subject, organizer, requiredAttendees, start, end, user }) {
            if (meetingType === MEETING_TYPES.OOTO) {
                // Create OOTO meetings directly here instead of calling a separate method
                return new Promise(async (resolve, reject) => {
                    try {
                        const emailContent = await getBody();

                        // If emailContent is null, the user clicked Cancel
                        if (emailContent === null) {
                            resolve(null);
                            return;
                        }
                        const subject = `OOTO | ${organizer} | ${getFormattedDateRange(start, end)}`;

                        await Promise.all([
                            // Team meeting (Free)
                            new Promise((resolve, reject) => {
                                GM.xmlHttpRequest({
                                    method: 'POST',
                                    url: `${this.endpoint}/owa/service.svc`,
                                    headers: {
                                        action: 'CreateItem',
                                        'content-type': 'application/json; charset=UTF-8',
                                        'x-owa-actionname': 'CreateCalendarItemAction',
                                        'x-owa-canary': this.getToken(),
                                        ...(this.getJwt() && { 'authorization': this.getJwt() })
                                    },
                                    data: JSON.stringify({
                                        Header: {
                                            RequestServerVersion: 'Exchange2013',
                                            TimeZoneContext: {
                                                TimeZoneDefinition: { Id: 'UTC' }
                                            }
                                        },
                                        Body: {
                                            Items: [{
                                                __type: 'CalendarItem:#Exchange',
                                                Subject: subject,
                                                Body: {
                                                    BodyType: 'HTML',
                                                    Value: emailContent
                                                },
                                                Sensitivity: 'Normal',
                                                IsResponseRequested: false,
                                                Start: start.toISOString(),
                                                End: end.toISOString(),
                                                FreeBusyType: 'Free',
                                                RequiredAttendees: [{
                                                    __type: 'AttendeeType:#Exchange',
                                                    Mailbox: {
                                                        EmailAddress: 'all-microsites@amazon.com',
                                                        RoutingType: 'SMTP',
                                                        MailboxType: 'Mailbox',
                                                        OriginalDisplayName: 'all-microsites@amazon.com'
                                                    }
                                                }]
                                            }],
                                            SendMeetingInvitations: 'SendToAllAndSaveCopy'
                                        }
                                    }),
                                    onerror: this.onajaxerror(reject),
                                    onload: response => {
                                        if (response.status === 200) {
                                            resolve();
                                        } else {
                                            this.onajaxerror(reject)(response);
                                        }
                                    }
                                });
                            }),

                            // Self meeting (OOF)
                            new Promise((resolve, reject) => {
                                GM.xmlHttpRequest({
                                    method: 'POST',
                                    url: `${endpoint}/owa/service.svc`,
                                    headers: {
                                        action: 'CreateItem',
                                        'content-type': 'application/json; charset=UTF-8',
                                        'x-owa-actionname': 'CreateCalendarItemAction',
                                        'x-owa-canary': this.getToken(),
                                        ...(this.getJwt() && { 'authorization': this.getJwt() })
                                    },
                                    data: JSON.stringify({
                                        Header: {
                                            RequestServerVersion: 'Exchange2013',
                                            TimeZoneContext: {
                                                TimeZoneDefinition: { Id: 'UTC' }
                                            }
                                        },
                                        Body: {
                                            Items: [{
                                                __type: 'CalendarItem:#Exchange',
                                                Subject: subject,
                                                Body: {
                                                    BodyType: 'HTML',
                                                    Value: emailContent
                                                },
                                                Sensitivity: 'Normal',
                                                IsResponseRequested: false,
                                                Start: start.toISOString(),
                                                End: end.toISOString(),
                                                FreeBusyType: 'OOF',
                                                RequiredAttendees: [{
                                                    __type: 'AttendeeType:#Exchange',
                                                    Mailbox: {
                                                        EmailAddress: organizer,
                                                        RoutingType: 'SMTP',
                                                        MailboxType: 'Mailbox',
                                                        OriginalDisplayName: organizer
                                                    }
                                                }]
                                            }],
                                            SendMeetingInvitations: 'SendToAllAndSaveCopy'
                                        }
                                    }),
                                    onerror: this.onajaxerror(reject),
                                    onload: response => {
                                        if (response.status === 200) {
                                            resolve();
                                        } else {
                                            this.onajaxerror(reject)(response);
                                        }
                                    }
                                });
                            })
                        ]);
                        resolve(true);
                    } catch (error) {
                        reject(error);
                    }
                });
            } else if (meetingType === MEETING_TYPES.QUICK) {
                // Create Quick meeting
                return new Promise(async (resolve, reject) => {
                    try {
                        const emailContent = await getQuickMeetingBody(user);

                        // If emailContent is null, the user clicked Cancel
                        if (emailContent === null) {
                            resolve(null);
                            return;
                        }

                        GM.xmlHttpRequest({
                            method: 'POST',
                            url: `${endpoint}/owa/service.svc`,
                            headers: {
                                action: 'CreateItem',
                                'content-type': 'application/json; charset=UTF-8',
                                'x-owa-actionname': 'CreateCalendarItemAction',
                                'x-owa-canary': this.getToken(),
                                ...(this.getJwt() && { 'authorization': this.getJwt() })
                            },
                            data: JSON.stringify({
                                Header: {
                                    RequestServerVersion: 'Exchange2013',
                                    TimeZoneContext: {
                                        TimeZoneDefinition: { Id: 'UTC' }
                                    }
                                },
                                Body: {
                                    Items: [{
                                        __type: 'CalendarItem:#Exchange',
                                        Subject: subject,
                                        Body: {
                                            BodyType: 'HTML',
                                            Value: emailContent
                                        },
                                        Sensitivity: 'Normal',
                                        IsResponseRequested: true,
                                        Start: start.toISOString(),
                                        End: end.toISOString(),
                                        FreeBusyType: 'Busy',
                                        RequiredAttendees: requiredAttendees.map(attendee => ({
                                            __type: 'AttendeeType:#Exchange',
                                            Mailbox: {
                                                EmailAddress: attendee.email,
                                                RoutingType: 'SMTP',
                                                MailboxType: 'Mailbox',
                                                OriginalDisplayName: attendee.email
                                            }
                                        }))
                                    }],
                                    SendMeetingInvitations: 'SendToAllAndSaveCopy'
                                }
                            }),
                            onerror: this.onajaxerror(reject),
                            onload: response => {
                                if (response.status === 200) {
                                    resolve(true);
                                } else {
                                    this.onajaxerror(reject)(response);
                                }
                            }
                        });
                    } catch (error) {
                        reject(error);
                    }
                });
            }
        },

        getJwt () {
            return GM_getValue('outlookJwtToken', '');
        },

        getToken () {
            return localStorage.phonetoolCalendarOutlookToken;
        },

        updateToken (response) {
            const match = response.responseHeaders.match(/x-owa-canary=(.*?);/i);
            if (match) {
                localStorage.phonetoolCalendarOutlookToken = match[1];
            }
        }
    })

    const EXCHANGE_DATA_PROVIDER = buildDataProvider('exchange', 'https://ballard.amazon.com');
    const M365_DATA_PROVIDER = buildDataProvider('m365', 'https://outlook.office.com');

    function showEmailContentDialog(emailContent, callback) {
        // Convert HTML content to plain text
        const plainTextContent = emailContent
        .replace(/<br>/g, '\n')
        .replace(/<[^>]*>/g, '')
        .replace(/&nbsp;/g, ' ')
        .trim();

        // Create the dialog container
        const dialogHTML = `
        <div id="email-content-dialog" style="
            display: none;
            position: fixed;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
            padding: 20px;
            background: white;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0, 0, 0, 0.1);
            z-index: 1001;
            width: 600px;">
            <h3 style="margin-top: 0; color: #232f3e;">Edit Email Content</h3>
            <textarea id="email-content-editor" style="
                width: 100%;
                height: 300px;
                margin: 10px 0;
                padding: 10px;
                border: 1px solid #ccc;
                border-radius: 4px;
                resize: vertical;
                font-family: Arial, sans-serif;
                font-size: 14px;
                line-height: 1.4;">${plainTextContent}</textarea>
            <div style="text-align: right; margin-top: 10px;">
                <button id="cancel-edit" style="
                    margin-right: 10px;
                    padding: 8px 16px;
                    background: #ff6b6b;
                    color: white;
                    border: none;
                    border-radius: 4px;
                    cursor: pointer;">Cancel</button>
                <button id="send-without-edit" style="
                    margin-right: 10px;
                    padding: 8px 16px;
                    background: #f0f0f0;
                    border: none;
                    border-radius: 4px;
                    cursor: pointer;">Send Without Edit</button>
                <button id="save-and-send" style="
                    padding: 8px 16px;
                    background: #0066cc;
                    color: white;
                    border: none;
                    border-radius: 4px;
                    cursor: pointer;">Save & Send</button>
            </div>
        </div>
        <div id="email-content-overlay" style="
            display: none;
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.5);
            z-index: 1000;">
        </div>`;

        // Add the dialog to the document
        document.body.insertAdjacentHTML('beforeend', dialogHTML);

        const dialog = document.getElementById('email-content-dialog');
        const overlay = document.getElementById('email-content-overlay');
        const editor = document.getElementById('email-content-editor');

        // Show the dialog
        dialog.style.display = 'block';
        overlay.style.display = 'block';

        // Function to clean up and close the dialog
        const closeDialog = () => {
            dialog.remove();
            overlay.remove();
        };

        // Handle buttons
        document.getElementById('cancel-edit').addEventListener('click', () => {
            closeDialog();
            callback(null); // Send null to indicate cancellation
        });

        document.getElementById('send-without-edit').addEventListener('click', () => {
            closeDialog();
            callback(emailContent); // Send original HTML content
        });

        document.getElementById('save-and-send').addEventListener('click', () => {
            const editedContent = editor.value
            .split('\n')
            .map(line => line.trim())
            .join('<br>'); // Convert newlines back to <br> tags
            closeDialog();
            callback(editedContent);
        });
    }

    function getBody() {
        return new Promise((resolve) => {
            const supervisorXPath = "/html/body/div[2]/div[1]/div/div/form/div[1]/div[2]/div[4]/div[2]/p/a";
            const supervisorElement = document.evaluate(supervisorXPath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
            const supervisorEmail = supervisorElement ? supervisorElement.href.replace('mailto:', '') : 'supervisor@amazon.com';

            const nameXPath = "/html/body/div[2]/div[7]/div/div[1]/div[2]/div[2]/div/div[1]/div[1]/div[1]/div";
            const nameElement = document.evaluate(nameXPath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
            let yourName = nameElement ? nameElement.textContent.trim() : '';

            if (yourName.includes(',')) {
                yourName = yourName.split(',')[0].trim();
            }

            // Add <br> tags for line breaks in the HTML content
            const emailContent = `Hello Team,<br><br>
I will be out of office on the scheduled dates with no access to outlook, slack and chime.<br><br>
For project related queries, please reach out to ${supervisorEmail}.<br><br>
Regards,<br>
${yourName}<br>
=====================================================`;

            showEmailContentDialog(emailContent, (content) => {
                resolve(content);
            });
        });
    }

    function getQuickMeetingBody(user) {
        return new Promise((resolve) => {
            const emailContent = `==============Quick Meeting Information==============

You have been invited to a quick meeting with ${user}.

Join via Chime clients (auto-call): Chime will call you when the meeting starts.

This is a quick meeting scheduled through Phonetool Calendar.

=====================================================`;

            showEmailContentDialog(emailContent, (content) => {
                // Convert plain text back to HTML with <br> tags
                resolve(content);
            });
        });
    }

    let currentUser;
    function getCurrentUser () {
        if (!currentUser) {
            currentUser = JSON.parse(document.querySelector('[data-react-class=NavBar]').dataset.reactProps).currentUser;
        }
        return currentUser;
    }

    let targetUser;
    function getTargetUser () {
        if (!targetUser) {
            const userDetails = document.querySelector('[data-react-class="UserDetails"]');
            targetUser = userDetails && JSON.parse(userDetails.dataset.reactProps).targetUser;
        }
        return targetUser;
    }

    function isSundayToThursdayWeek() {
        return getTargetUser().targetUserBuilding.match(/^(TLV|HFA|AMM|CAI|JED|RUH|DMM|AHB|ELQ|HOF|ALY|LXR|KWI)/);
    }

    const Days = {
        Monday: 1,
        Tuesday: 2,
        Wednesday: 3,
        Thursday: 4,
        Friday: 5,
        Saturday: 6,
        Sunday: 0
    };

    function getWeekend() {
        return isSundayToThursdayWeek() ? [Days.Friday, Days.Saturday] : [Days.Saturday, Days.Sunday];
    }

    function getBusinessHours () {
        const commonWeekdays = [Days.Monday, Days.Tuesday, Days.Wednesday, Days.Thursday];
        const daysOfWeek = isSundayToThursdayWeek() ? [Days.Sunday, ...commonWeekdays] : [...commonWeekdays, Days.Friday]; // Days of week

        const UTCOffset = getTargetUser().targetUserUTCOffset;
        const offsetInMinutes = moment().utcOffset() - moment().utcOffset(UTCOffset).utcOffset();

        const startTime = moment().hours(9).minutes(offsetInMinutes).format('HH:mm');
        const endTime = moment().hours(18).minutes(offsetInMinutes).format('HH:mm');

        // This can be simplified when https://github.com/fullcalendar/fullcalendar/issues/4440 is resolved
        return startTime < endTime
            ? { daysOfWeek, startTime, endTime }
        : [{ daysOfWeek, startTime: '00:00', endTime }, { daysOfWeek, startTime, endTime: '24:00' }];
    }

    async function getCredentials(dataProvider) {
        const user = getTargetUser().targetUserLogin;
        const currentUser = getCurrentUser();

        const email = dataProvider.fetchEmail(user);
        const currentEmail = user === currentUser ? email : dataProvider.fetchEmail(currentUser);
        return Promise.all([dataProvider, user, currentUser, email, currentEmail]);
    }

    function shouldShowMyEvents() {
        const val = localStorage.phonetoolCalendarShowMyEvents;
        return val && JSON.parse(val);
    }

    function shouldShowHireEvents() {
        const val = localStorage.phonetoolCalendarShowHireEvents;
        return val && JSON.parse(val);
    }

    function convertEvents(events) {
        for(let event of events) {
            event.end = event.end._d;
            event.start = event.start._d;
            if (event.title === 'tentative' || event.className.includes('tentative')) {
                event.textColor = '#555';
            }
        }
        return events;
    }

    function createOOTOForm(container) {
        const formHTML = `
<div id="ooto-form" style="
    display: none;
    position: fixed;
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%);
    padding: 25px;
    background: linear-gradient(135deg, #eaeded, #eaeded);
    border-radius: 15px;
    border: none;
    box-shadow: 0 10px 30px rgba(0, 0, 0, 0.5);
    z-index: 1000;
    font-family: 'Poppins', sans-serif;
    width: 400px;
    color: #f1f1f1;">
    <h2 style="
        margin: 0 0 15px 0;
        font-size: 22px;
        color: #232f3e;
        text-align: center;
        text-shadow: 0 2px 5px rgba(0, 0, 0, 0.4);">
        Set Out of Office
    </h2>
    <label style="font-size: 14px; color: #232f3e; margin-bottom: 5px; display: block;">Date & Time:</label>
    <input type="datetime-local" id="ooto-start-datetime" style="
        width: 100%;
        padding: 10px;
        background: #2e2e3e;
        border: none;
        border-radius: 8px;
        font-size: 14px;
        margin-bottom: 10px;
        color: #ffffff;
        outline: none;
        box-shadow: inset 0 2px 4px rgba(0, 0, 0, 0.3);" />
    <input type="datetime-local" id="ooto-end-datetime" style="
        width: 100%;
        padding: 10px;
        background: #2e2e3e;
        border: none;
        border-radius: 8px;
        font-size: 14px;
        color: #ffffff;
        outline: none;
        box-shadow: inset 0 2px 4px rgba(0, 0, 0, 0.3);" />

    <div style="margin-top: 20px; text-align: center;">
        <button id="submit-ooto" style="
            background: linear-gradient(135deg, #1db954, #1ed760);
            color: white;
            padding: 10px 20px;
            border: none;
            border-radius: 10px;
            font-size: 14px;
            cursor: pointer;
            transition: transform 0.2s, box-shadow 0.2s;
            box-shadow: 0 4px 10px rgba(0, 0, 0, 0.3);">
            Submit
        </button>
        <button id="cancel-ooto" style="
            background: linear-gradient(135deg, #ff6a6a, #ff3e3e);
            color: white;
            padding: 10px 20px;
            border: none;
            border-radius: 10px;
            font-size: 14px;
            margin-left: 10px;
            cursor: pointer;
            transition: transform 0.2s, box-shadow 0.2s;
            box-shadow: 0 4px 10px rgba(0, 0, 0, 0.3);">
            Cancel
        </button>
    </div>
</div>
<div id="ooto-overlay" style="display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0, 0, 0, 0.7); z-index: 999;"></div>
<button id="open-ooto-form" style="
    background: linear-gradient(135deg, #049796, #049796);
    color: white;
    padding: 4px 10px;
    border: none;
    border-radius: 10px;
    font-size: 10px;
    cursor: pointer;
    transition: transform 0.2s, box-shadow 0.2s;
    box-shadow: 0 4px 10px rgba(0, 0, 0, 0.3);">
    Set Out of Office
</button>
`;
        container.innerHTML = formHTML;

        const form = document.getElementById('ooto-form');
        const overlay = document.getElementById('ooto-overlay');

        document.getElementById('open-ooto-form').addEventListener('click', () => {
            form.style.display = 'block';
            overlay.style.display = 'block';
        });

        document.getElementById('cancel-ooto').addEventListener('click', () => {
            form.style.display = 'none';
            overlay.style.display = 'none';
        });

        async function checkUserMailbox(dataProvider, email) {
            try {
                await dataProvider.getAvailability([email], new Date(), new Date(Date.now() + 86400000), email, true);
                return true;
            } catch (e) {
                console.error('Error checking mailbox:', e);
                return false;
            }
        }

        document.getElementById('submit-ooto').addEventListener('click', async () => {
            try {
                const startDatetime = document.getElementById('ooto-start-datetime').value;
                const endDatetime = document.getElementById('ooto-end-datetime').value;
                const username = getCurrentUser();
                const formattedStart = moment(new Date(startDatetime)).format('YYYY-MM-DD');
                const formattedEnd = moment(new Date(endDatetime)).format('YYYY-MM-DD');
                const subject = `OOTO | ${username} | (${formattedStart} - ${formattedEnd})`;

                if (!startDatetime || !endDatetime) {
                    alert('Please fill in all required fields.');
                    return;
                }

                const start = new Date(startDatetime);
                const end = new Date(endDatetime);

                // Initialize the data provider
                const [dataProvider] = await initCredentials();
                dataProvider.endpoint = 'https://outlook.office.com';

                await dataProvider.createMeeting({
                    meetingType: MEETING_TYPES.OOTO,
                    subject,
                    organizer: username,
                    start,
                    end
                });

                alert('Out of Office set successfully!');
                document.getElementById('ooto-form').style.display = 'none';
                document.getElementById('ooto-overlay').style.display = 'none';
            } catch (err) {
                console.error(err);
                alert('Error setting Out of Office. Please try again later.');
            }
        });
    }

    /**
     * Allows dragging the widget to the widgets area and back. The position is saved in the localStorage
     *
     * This is copied and pasted from `enable_widget_movements` method with few overrides
     */
    function enableContainerDragging (container) {
        const $slots = $('#widgets1, #widgets2, #calendar-container');
        if (!$slots.sortable) {
            // For unknown reason, some users have an issue when $(...).sortable doesn't work which breaks the tool
            // This workaround disables container dragging, however, the main functionality remains unimpared
            console.error('$.fn.sortable method is not available');
            return;
        }
        $slots.sortable({
            handle: '.widget-move-handle',
            connectWith: '#widgets1, #widgets2, #calendar-container',
            start: function (e, ui) {
                ui.placeholder.height(ui.helper.outerHeight());
                $('#calendar-container').css('min-height', '100px');
            },
            placeholder: 'widget-placeholder',
            forceHelperSize: true,
            forcePlaceHolderSize: true,
            tolerance: 'pointer', // Overriden
            over: unsafeWindow.keep_equal_heights,
            update: function () {
                unsafeWindow.update_positions();
                // Overriden behavior
                // Every time after resorting, we calculate and save container position
                const parentElement = container.parentElement;
                const parentSelector = '#' + parentElement.id;
                const itemIndex = [...parentElement.children].findIndex(_ => _ == container);
                localStorage.phonetoolCalendarContainer = [parentSelector, itemIndex].join('%');

                if (parentElement.id == 'calendar-container') {
                    container.classList.remove('well');
                    document.querySelector('.SecondaryDetails').classList.add('calendar-atf-container');
                } else {
                    container.classList.add('well');
                    document.querySelector('.SecondaryDetails').classList.remove('calendar-atf-container');
                }
            },
            sort: function (e, ui) {
                // Overriden behavior
                // Prevent dropping the default widgets to above the fold area
                if (ui.item[0] != container && ui.placeholder.parent()[0].id == 'calendar-container') {
                    return false;
                }
            },
            stop: function () {
                // Overriden behavior
                $('#calendar-container').css('min-height', '0');
            }
        });
    }

    function initCss() {
        const css = document.createElement('link');
        const externalCdn = "https://cdn.jsdelivr.net/npm/";
        const fullCalendarCdn = "fullcalendar@5.11.5/main.min.css";
        css.href = externalCdn + fullCalendarCdn;
        css.rel = 'stylesheet';
        document.head.appendChild(css);

        document.head.appendChild(document.createElement('style')).textContent = `
            .SecondaryDetails.calendar-atf-container {
                align-items: stretch;
            }
            .calendar-atf-container .UserLinks, .calendar-atf-container .ResolverRow { white-space: nowrap }

            .calendar-atf-container .SharePassion {
                flex-grow: 1;
                max-width: 1000px;
            }

            .calendar-widget .widget-move-handle {
                position: relative;
                height: 17px;
                padding-left: 20px;
                background: url('data:image/svg+xml;utf8,<svg viewBox="0 0 25 30" xmlns="http://www.w3.org/2000/svg" fill="grey"><circle cx="5" cy="5" r="3"/><circle cx="18" cy="5" r="3"/><circle cx="5" cy="15" r="3"/><circle cx="18" cy="15" r="3"/><circle cx="5" cy="25" r="3"/><circle cx="18" cy="25" r="3"/></svg>') no-repeat;
            }

            .script-warning { display: none; }
            .error-unauthorized .script-warning { display: block; font-weight: bold; }
            .error-unauthorized .calendar-my-events-label { display: none; }
            .error-unauthorized .calendar-hire-events-label { display: none; }

            .calendar-tentative, .calendar-tentative:hover {
              background: repeating-linear-gradient(-48deg, white, white 1px, #99c8e9 3px, #99c8e9 3px, white 6px);
              color: #555;
            }
            .calendar-out-of-office { border-color: #800080; background: #800080; }
            .calendar-no-data { border-color: #888; background: #888; }
            .calendar-hire { border-color: #cf4f13; background: #cf4f13; }

            .calendar-my-busy { border-color: #86ac39; background: #86ac39; }
            .calendar-my-tentative, .calendar-my-tentative:hover {
              border-color: #86ac39;
              background: repeating-linear-gradient(-48deg, white, white 1px, #cfea9a 3px, #cfea9a 3px, white 6px);
              color: #555;
            }
            .calendar-my-out-of-office { border-color: #af6aaf; background: #af6aaf; }
            .calendar-my-hire { border-color: #cf4f13; background: #cf4f13; }

            .fc-event.fc-v-event { cursor: default; }

            .fc-nonbusiness { background: #bbb; } /* Darker background for non-business hours */

            .calendar-container { margin-top: 7px; display: inline-flex; width: 100%; min-width: 540px; height: 400px; }
            .calendar-my-events-label { float: right; }
            .calendar-my-events-checkbox { transform: translateX(-5px); vertical-align: baseline; }

            .calendar-hire-events-label { float: right; }
            .calendar-hire-events-checkbox { transform: translateX(-5px); vertical-align: baseline; }

            .calendar-container + hr { display: none; }
            .SharePassion .calendar-container + hr { display: block; }
        `;
    }

    async function initCredentials() {
        // Amazon is slowly migrating to M365, some users will be on Exchange others will be on M365
        for (const dp of [M365_DATA_PROVIDER, EXCHANGE_DATA_PROVIDER]) {
            try {
                console.info(`Trying to connect to ${dp.name}`);
                return await getCredentials(dp);
            } catch (e) {
                console.warn(`Failed to connect to ${dp.name}`, e);
            }
        }
        throw 'Cannot get credentials from any data provider';
    }

    async function phonetoolScript () {
        if (!getTargetUser() || !getTargetUser().targetUserActive) {
            return; // The user is inactive or doesn't exist
        }

        const credentials = initCredentials();

        initCss();

        await Promise.all([
            'https://cdn.jsdelivr.net/npm/fullcalendar@5.11.5/main.min.js',
            'https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.22.2/moment.min.js'
        ].map(src => new Promise(resolve => {
            const script = document.createElement('script');
            script.src = src;
            script.onload = resolve;
            document.body.appendChild(script);
        })));

        const container = document.createElement('div');
        container.className = 'calendar-widget';
        const droppableContainer = document.createElement('div');
        droppableContainer.id = 'calendar-container';
        droppableContainer.style = 'min-width: 50px';
        document.querySelector('.SharePassion').prepend(droppableContainer);
        const [parentSelector, containerIndex] = (localStorage.phonetoolCalendarContainer || '#calendar-container%0').split('%');
        const parentContainer = document.querySelector(parentSelector);
        parentContainer.insertBefore(container, parentContainer.children[containerIndex]);
        let hasError = false;

        // Awaiting the credentials __before__ actual changes to html in order to fail gracefully
        let dataProvider, user, currentUser, email, currentEmail;
        try {
            [dataProvider, user, currentUser, email, currentEmail] = await credentials;
            console.info("Successfully connected to: ", dataProvider.name);
        } catch(e) {
            // high chances of missing Microsoft Exchange authentication token: show the error message
            console.error("Unauthorized ", e);
            container.classList.add('error-unauthorized');
            hasError = true;
        }

        const showMyEventsButton = user !== currentUser;
        const isSelfView = user == currentUser;
        const myEventsButton = showMyEventsButton ? `<label class="calendar-my-events-label"><input class="calendar-my-events-checkbox" type="checkbox">Display my events</label>` : '';
        const hireEventsButton = isSelfView ? `<label class="calendar-hire-events-label"><input title="Hire events will be highlighted in a separate color" class="calendar-hire-events-checkbox" type="checkbox">Highlight hire events</label>` : '';

        container.innerHTML = `
          <div class="widget-move-handle">
            ${myEventsButton} ${hireEventsButton}
            <div style="position: absolute; left: 50%; transform: translateX(-50%);">
              <a href="https://w.amazon.com/bin/view/Scrat/Tools/PhonetoolCalendar/#HChangelog">
                Calendar Wiki
                ${localStorage.phonetoolCalendarVersion == GM_info.script.version ? '' : '<b>(see what is new in version ' + GM_info.script.version + ')</b>'}
              </a>
            </div>
          </div>
          <p class="script-warning">
            Unable to fetch calendar data, usually due to missing authentication token.<br>
            If your inbox has migrated to Microsoft 365:<br>
            Try to login on <a href="${M365_DATA_PROVIDER.endpoint}/owa">${M365_DATA_PROVIDER.name}</a> and then refresh this page.<br>
            <br>
            If your inbox is still on Exchange:<br>
            Try to login on <a href="${EXCHANGE_DATA_PROVIDER.endpoint}/owa">${EXCHANGE_DATA_PROVIDER.name}</a> and then refresh this page.<br>
            <br>
            For more information please check out <a href="https://w.amazon.com/bin/view/Scrat/Tools/PhonetoolCalendar#HTroubleshooting2FFAQ">the Calendar's wiki page</a> that is being kept updated.<br>
            <br>
           </p>
          <div id="cc-content" class="calendar-container"></div>
          <hr>
        `;

        if (hasError) {
            document.querySelector('#cc-content').remove();
        }

        if (parentContainer.classList.contains('widgets')) {
            container.classList.add('well');
        } else {
            document.querySelector('.SecondaryDetails').classList.add('calendar-atf-container');
        }

        if (showMyEventsButton) {
            container.querySelector('.calendar-my-events-checkbox').checked = shouldShowMyEvents();
        }

        if (isSelfView) {
            container.querySelector('.calendar-hire-events-checkbox').checked = shouldShowHireEvents();
        }

        localStorage.phonetoolCalendarVersion = GM_info.script.version;

        enableContainerDragging(container);

        // Stop rendering the calendar, if the user is missing (due to some load error)
        if (!user) {
            return;
        }

        const calendarEl = container.querySelector('.calendar-container');

        if (showMyEventsButton) {
            $(container.querySelector('.calendar-my-events-checkbox')).on('change', _ => {
                localStorage.phonetoolCalendarShowMyEvents = _.target.checked;
                calendar.destroy();
                calendar = new FullCalendar.Calendar(calendarEl, calendarParams);
                calendar.render();
            });
        }

        if (isSelfView) {
            $(container.querySelector('.calendar-hire-events-checkbox')).on('change', _ => {
                localStorage.phonetoolCalendarShowHireEvents = _.target.checked;
                calendar.destroy();
                calendar = new FullCalendar.Calendar(calendarEl, calendarParams);
                calendar.render();
            });
        }

        const subjectText = user === currentUser
        ? `This will create an Out of the Office calendar entry`
            : `This will create a quick meeting with @${user}`;

        const calendarParams = {
            allDaySlot: false,
            headerToolbar: {
                left: 'prev,next today',
                center: 'title',
                right: 'dayGridMonth,timeGridWeek,timeGridDay,listWeek'
            },
            businessHours: getBusinessHours(),
            nowIndicator: true,
            hiddenDays: getWeekend(),
            initialView: 'timeGridWeek',
            navLinks: true, // can click day/week names to navigate views
            // editable: true,
            scrollTime: '09:00:00',
            selectable: true,
            selectMirror: true,
            select: function(selectionInfo) {
                const [start, end] = [selectionInfo.start, selectionInfo.end];
                if (start < moment()._d) {
                    alert('Cannot create a meeting in the past');
                } else {
                    const subject = prompt(`${subjectText}. If agree, please enter the subject.`);
                    if (subject !== null) {
                        const requiredAttendees = [{ email: 'meet@chime.aws' }];
                        if (email != currentEmail) {
                            requiredAttendees.push({ email });
                        }
                        dataProvider.createMeeting({ subject, organizer: currentEmail, requiredAttendees, start, end })
                            .then(_ => {
                            alert(`The meeting "${subject}" is successfully created!`);
                            calendar.addEvent({ start, end });
                        })
                            .catch(_ => {
                            console.error(_);
                            alert('Error');
                        });
                    }
                }
                calendar.unselect();
            },
            dayMaxEventRows: true, // allow "more" link when too many events
            events: function (fetchInfo, successCallback, failureCallback) {
                const mailboxes = [email];
                if (email != currentEmail && shouldShowMyEvents()) {
                    mailboxes.push(currentEmail);
                }

                dataProvider.getAvailability(mailboxes, fetchInfo.start, fetchInfo.end, mailboxes[1], isSelfView).then(events => {
                    const convertedEvents = convertEvents(events);
                    successCallback([].concat(...convertedEvents));
                }).catch(e => {
                    console.error('Error fetching from ', dataProvider.name, ': ', e)
                });
            }
        };

        var calendar = new FullCalendar.Calendar(calendarEl, calendarParams);
        calendar.render();

        let resizeObserver = new ResizeObserver(() => {
            calendar.updateSize()
        });
        resizeObserver.observe(container);
    }

    const container = document.createElement('div');
    container.className = 'ooto-widget';
    document.querySelector('.SharePassion').prepend(container);
    createOOTOForm(container);

    phonetoolScript();

    // Token interception for outlook.office.com
    if (window.location.hostname === 'outlook.office.com') {
        // Intercept fetch requests
        const originalFetch = unsafeWindow.fetch || window.fetch;
        const interceptedFetch = function(...args) {
            const [resource, config] = args;
            const url = typeof resource === 'string' ? resource : resource.url;

            if (url && url.includes('/owa/service.svc') && config && config.headers) {
                const headers = config.headers;
                let authHeader = null;

                // Headers can be a Headers object, plain object, or array
                if (headers instanceof Headers) {
                    authHeader = headers.get('authorization');
                } else if (Array.isArray(headers)) {
                    const authEntry = headers.find(([key]) => key.toLowerCase() === 'authorization');
                    authHeader = authEntry ? authEntry[1] : null;
                } else if (typeof headers === 'object') {
                    authHeader = headers.authorization || headers.Authorization;
                }

                if (authHeader) {
                    GM_setValue('outlookJwtToken', authHeader);
                }
            }

            return originalFetch.apply(unsafeWindow || window, args);
        };

        // Replace fetch
        if (unsafeWindow) unsafeWindow.fetch = interceptedFetch;
        window.fetch = interceptedFetch;
    }

})();
