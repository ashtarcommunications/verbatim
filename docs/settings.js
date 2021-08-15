function setSetting(setting, value) {
    var prop = PropertiesService.getUserProperties();
    prop.setProperty(setting, value);
    return true;
}

function getSetting(setting) {
    var prop = PropertiesService.getUserProperties();
    return prop.getProperty(setting);
}
