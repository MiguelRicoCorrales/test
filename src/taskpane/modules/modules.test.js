import { SatellitesQuery, ParametersQuery, LevelsQuery, BuildServiceAddress } from './H4OSvc';
const DEFAULT_HOST = "demo-swarm"
const DEFAULT_PORT = 7000


describe("Calls to archiva", () => {

    it('Satellites info is read correctly', () => {

        const satellites = []
        const headers = {
            'X-Requested-With': 'XMLHttpRequest'
        };

        fetch(SatellitesQuery(DEFAULT_HOST, DEFAULT_PORT), { headers })
            .then(response => response.text())
            .then(html => {
                const tempElement = document.createElement('div');
                tempElement.innerHTML = html;
                const satelliteRows = tempElement.querySelectorAll('table tr');
                const sats = Array.from(satelliteRows).map(row => row.querySelector('td').textContent);
                sats.forEach(sat => {
                    satellites.push(sat)
                });

                console.log("satellites", satellites)

            })
        console.log("satellites out", satellites)

    });


})
