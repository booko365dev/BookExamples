//gavdcodebegin 008
//Replaced by TZGQ
(function () {
  'use strict';

    microsoftTeams.initialize();

    microsoftTeams.getContext(function (context) {
        if (context && context.theme) {
            setTheme(context.theme);
        }
    });

    microsoftTeams.registerOnThemeChangeHandler(function (theme) {
        setTheme(theme);
    });

    microsoftTeams.settings.registerOnSaveHandler(function (saveEvent) {
        microsoftTeams.settings.setSettings({
            contentUrl: createTabUrl(), // Mandatory parameter
            entityId: createTabUrl() // Mandatory parameter
        });

        saveEvent.notifySuccess();
    });

    document.addEventListener('DOMContentLoaded', function () {
        var tabChoice = document.getElementById('tabChoice');
        if (tabChoice) {
            tabChoice.onchange = function () {
                var selectedTab = this[this.selectedIndex].value;
                microsoftTeams.settings.setValidityState(selectedTab ===
                    'GenerateGuid.aspx' || selectedTab === 'GenerateRandomString.aspx');
            };
    }
  });

    function setTheme(theme) {
        if (theme) {
            // Possible values for theme: 'default', 'light', 'dark' and 'contrast'
            document.body.className = 'theme-' + (theme === 'default' ? 'light' : theme);
        }
    }

    function createTabUrl() {
        var tabChoice = document.getElementById('tabChoice');
        var selectedTab = tabChoice[tabChoice.selectedIndex].value;

        return window.location.protocol + '//' + window.location.host + '/' + selectedTab;
    }
})();
//gavdcodeend 008

