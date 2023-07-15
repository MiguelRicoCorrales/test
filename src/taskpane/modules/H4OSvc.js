export const SatellitesQuery = (host, port) => {
  return BuildServiceAddress(host, port) + "/satellites";
};

export const ParametersQuery = (host, port, satellite) => {
  return BuildServiceAddress(host, port) + "/parameters?satellite=" + satellite;
};

export const LevelsQuery = (host, port, satellite) => {
  return BuildServiceAddress(host, port) + "/levels?satellite=" + satellite;
};

export const SamplesQuery = (host, port, formData) => {
  const targetUrl =
    "/samples" +
    "?satellite=" +
    encodeURIComponent(formData.satellite) +
    "&level=" +
    encodeURIComponent(formData.level) +
    "&parameters=" +
    encodeURIComponent(formData.listSelected) +
    "&start=" +
    encodeURIComponent(formData.start) +
    "&end=" +
    encodeURIComponent(formData.end);

  return BuildServiceAddress(host, port) + encodeURIComponent(targetUrl);
};

export const GetTMQuery = (host, port, satellite, level, parameter, timestamp) => {
  const query =
    `${BuildServiceAddress(host, port)}/gettm` +
    `?satellite=${satellite}` +
    `&level=${level}` +
    `&parameter=${encodeURIComponent(parameter)}` +
    `&timestamp=${timestamp.toISOString()}` +
    `&separator=${getDecimalSeparator()}`;

  return query;
};
export const BuildServiceAddress = (host, port) => {
  return "http://localhost:5000/proxy?url=http://" + host + ":" + port + "/archiva/rest/v1/H4O";
};
