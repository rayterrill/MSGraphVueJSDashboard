Vue.component('workday-breakdown', {
    props: ['token'], //token is an input to our component
    template: '<canvas id="chart-area"></canvas>',
    mounted() {
        this.getWorkdayBreakdown()
    },
    methods: {
        getWorkdayBreakdown() {
            var _this = this; //save a reference to this to allow us to us the parent this to update our data

            //get working hours for the signed-in user
            $.ajax({
                type: "GET",
                url: "https://graph.microsoft.com/v1.0/me/mailboxSettings/workingHours",
                headers: {
                    'Authorization': 'Bearer ' + _this.token,
                }
            }).done(function (data) {
                //calculate minutes in my day
                day = moment().format('YYYY-MM-DD');
                startDate = moment(day + 'T' + data.startTime);
                endDate = moment(day + 'T' + data.endTime);
                minutesInMyDay = endDate.diff(startDate, 'minutes');

                //build dates for getting meetings
                var start = moment().startOf('day').add(1, 'second').format();
                var end = moment().endOf('day').format();

                //get all your meetings for today
                $.ajax({
                    type: "GET",
                    url: "https://graph.microsoft.com/v1.0/me/calendar/calendarView?startDateTime=" + start + "&endDateTime=" + end,
                    headers: {
                        'Authorization': 'Bearer ' + _this.token,
                        'Prefer': 'outlook.timezone="Pacific Standard Time"'
                    }
                }).done(function (data) {
                    minutesInMeetings = 0;

                    if (data.value === undefined || data.value.length == 0) {
                        minutesNotInMeetings = minutesInMyDay;
                    } else {
                        //iterate over the results
                        for (d in data.value) {
                            var start = moment(data.value[d].start.dateTime);
                            var end = moment(data.value[d].end.dateTime);

                            meetingDuration = end.diff(start, 'minutes');

                            minutesInMeetings = minutesInMeetings + meetingDuration;
                            minutesNotInMeetings = minutesInMyDay - minutesInMeetings;
                        }
                    }

                    var config = {
                        type: 'pie',
                        data: {
                            datasets: [{
                                data: [
                                    minutesInMeetings,
                                    minutesNotInMeetings
                                ],
                                backgroundColor: [
                                    'rgb(255, 99, 132)',
                                    'rgb(75, 192, 192)',
                                ],
                                label: 'Dataset 1'
                            }],
                            labels: [
                                'Meetings',
                                'Not in Meetings',
                            ]
                        }
                    };

                    $(function() {
                        var ctx = document.getElementById('chart-area');
                        window.myPie = new Chart(ctx, config);
                    });
                }).fail(function() {
                    console.log('Error getting calendar data.');
                });

                //make the hidden section visible
                //$('#data').text(output).show();

                $('#chartdata').text('A breakdown of your workday:').show();
            }).fail(function() {
                console.log('Error getting mailbox settings!');
            });
        }
    }
})

//a vue component for our top10 people
Vue.component('top-people', {
    data: function() {
        return {
            imageOutput: null
        }
    },
    props: ['token'], //token is an input to our component
    template: '<div v-html="imageOutput"></div>', //use imageOutput in our template along with the v-html functionality to render the images
    mounted() {
        this.getPeople()
    },
    methods: {
        getPeople() {
            var _this = this; //save a reference to this to allow us to us the parent this to update our data
            var output = '';

            //get working hours for the signed-in user
            $.ajax({
                type: "GET",
                url: "https://graph.microsoft.com/v1.0/me/people/?top=10",
                headers: {
                    'Authorization': 'Bearer ' + _this.token,
                }
            }).done(async function (data) {
                var promises = [];

                for (d in data.value) {
                    //add the requests to a "promises" array so we can wait for them all to finish later
                    promises.push($.ajax({
                        type: "GET",
                        url: "https://graph.microsoft.com/v1.0/users/" + data.value[d].userPrincipalName + "/photos/48x48/$value",
                        headers: {
                            'Authorization': 'Bearer ' + _this.token,
                        },
                        xhr:function(){ // Seems like the only way to get access to the xhr object
                            var xhr = new XMLHttpRequest();
                            xhr.responseType= 'blob'
                            return xhr;
                        },
                    }).done(function (data) {
                        var imageElm = document.createElement("img");
                        var reader = new FileReader();
                        reader.onload = function () {
                            // Add the base64 image to the src attribute
                            imageElm.src = reader.result;

                            //update the output variable with the image code
                            output = output + "<img src=\"" + imageElm.src + "\" />"
                        }
                        reader.readAsDataURL(data);
                    }).fail(function() {
                        console.log('Could not get photo');
                    }));
                }
                
                //wait for all the async calls to finish
                await Promise.all(promises.map(p => p.catch(() => undefined)));
                //update the vuejs data with the result
                _this.imageOutput = output;
            }).fail(function() {
                console.log('Error getting top 10 people!');
            });
        }
    }
});

var app = new Vue({
    el: '#app',
    data: {
      message: 'Hello Vue!',
      imageOutput: null,
      loading: false,
      token: null
    },
    created () {
        this.getToken()
    },
    methods: {
        getToken() {
            _this = this; //save a reference to the parent this
            // Enter Global Config Values & Instantiate ADAL AuthenticationContext
            window.config = {
                instance: 'https://login.microsoftonline.com/',
                tenant: TENANT,
                clientId: CLIENT_ID,
                postLogoutRedirectUri: window.location.origin,
                cacheLocation: 'localStorage' // enable this for IE, as sessionStorage does not work for localhost.
            };
            var authContext = new AuthenticationContext(config);

            // Get UI jQuery Objects
            var $panel = $(".panel-body");
            var $userDisplay = $(".app-user");
            var $signInButton = $(".app-login");
            var $signOutButton = $(".app-logout");
            var $errorMessage = $(".app-error");

            // Check For & Handle Redirect From AAD After Login
            var isCallback = authContext.isCallback(window.location.hash);
            authContext.handleWindowCallback();
            $errorMessage.html(authContext.getLoginError());

            if (isCallback && !authContext.getLoginError()) {
                window.location = authContext._getItem(authContext.CONSTANTS.STORAGE.LOGIN_REQUEST);
            }

            //if the user is not logged in, fire the azuread login logic
            if (!authContext.getCachedUser()) {
                authContext.config.redirectUri = [location.protocol, '//', location.host, location.pathname].join('');
                authContext.login();
            } else {
                // Check Login Status, Update UI
                var user = authContext.getCachedUser();
                if (user) {
                    $userDisplay.html(user.userName);
                    $userDisplay.show();
                    $signInButton.hide();
                    $signOutButton.show();
                    
                    //acquire token for ms graph. the service we're acquiring a token for should be the same service we call in the ajax request below
                    authContext.acquireToken('https://graph.microsoft.com', function (error, token) {
                        // Handle ADAL Error
                        if (error || !token) {
                            printErrorMessage('ADAL Error Occurred: ' + error);
                            return;
                        }
                    
                        _this.token = token; //update our data with the token
                    });
                } else {
                    $userDisplay.empty();
                    $userDisplay.hide();
                    $signInButton.show();
                    $signOutButton.hide();
                }
            }
        }
    }
  })