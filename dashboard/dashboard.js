fetch('dashboard_data.json')
            .then(res => res.json())
            .then(data => {
                // En-tête & KPIs
                document.getElementById('chantier-titre').innerText = data.chantier || "Suivi de Chantier";
                document.getElementById('export-date').innerText = "Dernier export : " + (data.date_export || "--");

                if (data.global) {
                    document.getElementById('kpi-global').innerText = (data.global.avancement_global_pct || 0) + "%";
                    document.getElementById('kpi-restants').innerText = (data.global.jours_restants !== null ? data.global.jours_restants + " j" : "--");
                }
                document.getElementById('kpi-retards').innerText = (data.retards || []).length;
                document.getElementById('kpi-entreprises').innerText = (data.charge || []).length;

                // 1. Graphique Barres (Avancement par lot)
                const ctxLots = document.getElementById('chartLots').getContext('2d');
                new Chart(ctxLots, {
                    type: 'bar',
                    data: {
                        labels: data.lots.map(l => "Lot " + l.numero_lot),
                        datasets: [{
                            label: '% Avancement',
                            data: data.lots.map(l => l.avancement_pct),
                            backgroundColor: '#0284c7'
                        }]
                    },
                    options: { scales: { y: { min: 0, max: 100 } } }
                });

                // 2. Graphique Donut (Charge entreprises)
                const ctxCharge = document.getElementById('chartCharge').getContext('2d');
                new Chart(ctxCharge, {
                    type: 'doughnut',
                    data: {
                        labels: data.charge.map(c => c.entreprise),
                        datasets: [{
                            data: data.charge.map(c => c.jours_planifies),
                            backgroundColor: ['#0284c7', '#38bdf8', '#f59e0b', '#ef4444', '#10b981', '#6366f1', '#ec4899']
                        }]
                    }
                });

                // 3. Remplissage du tableau des retards
                const tableBody = document.getElementById('retards-body');
                if (data.retards && data.retards.length > 0) {
                    data.retards.sort((a,b) => Number(a.numero_lot) - Number(b.numero_lot));
                    tableBody.innerHTML = data.retards.map(r => `
                        <tr>
                            <td>Lot ${r.numero_lot}</td>
                            <td><strong>${r.code_tache}</strong></td>
                            <td>${r.designation}</td>
                            <td>${r.entreprise}</td>
                            <td>${r.date_fin_prevue || '-'}</td>
                            <td><span class="badge-red">+${r.jours_de_retard} jours</span></td>
                            <td>${r.avancement_pct}%</td>
                            <td>${r.observations && r.observations !== 'None' ? r.observations : '-'}</td>
                        </tr>
                    `).join('');
                } else {
                    tableBody.innerHTML = `<tr><td colspan="8" style="text-align:center; color:#16a34a; padding: 20px;">Aucun retard constaté.</td></tr>`;
                }
            })
            .catch(err => {
                console.error("Erreur de chargement des données :", err);
            });