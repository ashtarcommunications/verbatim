function setProperty(property, value) {
    var properties = PropertiesService.getUserProperties();
    properties.setProperty(property, value);
    return true;
}

function getProperty(property) {
    var properties = PropertiesService.getUserProperties();
    return properties.getProperty(property);
}

const getProperties = () => {
    return PropertiesService.getUserProperties().getProperties();
}

const setProperties = (properties) => {
    return PropertiesService.getUserProperties().setProperties(properties);
}
