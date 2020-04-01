//gavdcodebegin 05
(function () {
    'use strict';

    microsoftTeams.initialize();

    microsoftTeams.getContext(function (context) {
        if (context && context.theme) {
            document.getElementById('lblContextInfo').innerText = context.loginHint;
            setTheme(context.theme);
        }
    });

    microsoftTeams.registerOnThemeChangeHandler(function (theme) {
        setTheme(theme);
    });

    function setTheme(theme) {
        if (theme) {
        // Possible values for theme: 'default', 'light', 'dark' and 'contrast'
        document.body.className = 'theme-' + (theme === 'default' ? 'light' : theme);
        }
    }
})();
//gavdcodeend 05
