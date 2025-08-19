  // Dados simulados mais realistas
    function getRandomAttitude() {
      return (Math.random() * 20 - 10).toFixed(1);
    }

    function getRandomDepth() {
      return (Math.random() * 50 + 5).toFixed(1);
    }

    function getRandomCygnus() {
      return (Math.random() * 30 + 70).toFixed(0);
    }

    function getRandomVoltage() {
      return (Math.random() * 4 + 10).toFixed(2); // Voltagem entre 10-14V
    }

    function getCpChannelVoltage() {
      return (Math.random() * 3.5).toFixed(2); // 0-3.5V para os canais CP
    }

    function getRandomReading() {
      return (Math.random() * 50 + 10).toFixed(1); // 10-60mm
    }

    // Variáveis para o gráfico
    let voltageChart = null;
    let chartData = {
      labels: [],
      datasets: [
        {
          label: 'CP CH1',
          data: [],
          borderColor: '#3b82f6',
          backgroundColor: 'rgba(59, 130, 246, 0.1)',
          borderWidth: 2,
          fill: false,
          tension: 0.3
        },
        {
          label: 'CP CH2',
          data: [],
          borderColor: '#22c55e',
          backgroundColor: 'rgba(34, 197, 94, 0.1)',
          borderWidth: 2,
          fill: false,
          tension: 0.3
        }
      ]
    };

    // Configurar o gráfico
    function initVoltageChart() {
      const ctx = document.getElementById('voltageChart');
      if (!ctx) return;

      voltageChart = new Chart(ctx, {
        type: 'line',
        data: chartData,
        options: {
          responsive: true,
          maintainAspectRatio: false,
          scales: {
            y: {
              beginAtZero: true,
              max: 3.5,
              grid: {
                color: 'rgba(255, 255, 255, 0.1)'
              },
              ticks: {
                color: 'rgba(255, 255, 255, 0.8)',
                stepSize: 0.5,
                callback: function(value) {
                  return value + 'V';
                }
              }
            },
            x: {
              grid: {
                color: 'rgba(255, 255, 255, 0.1)'
              },
              ticks: {
                color: 'rgba(255, 255, 255, 0.8)',
                maxTicksLimit: 10
              }
            }
          },
          plugins: {
            legend: {
              labels: {
                color: 'rgba(255, 255, 255, 0.9)',
                font: {
                  size: 14
                }
              }
            },
            tooltip: {
              backgroundColor: 'rgba(0, 0, 0, 0.8)',
              titleColor: 'white',
              bodyColor: 'white',
              borderColor: 'rgba(255, 255, 255, 0.2)',
              borderWidth: 1
            }
          },
          animation: {
            duration: 300
          }
        }
      });

      // Inicializar com alguns pontos
      for (let i = 0; i < 20; i++) {
        updateChartData();
      }
    }

    // Atualizar dados do gráfico
    function updateChartData() {
      if (!voltageChart) return;

      const now = new Date();
      const timeLabel = now.toLocaleTimeString('pt-BR', { 
        hour: '2-digit', 
        minute: '2-digit', 
        second: '2-digit' 
      });

      const ch1Value = parseFloat(getCpChannelVoltage());
      const ch2Value = parseFloat(getCpChannelVoltage());

      // Adicionar novos dados
      chartData.labels.push(timeLabel);
      chartData.datasets[0].data.push(ch1Value);
      chartData.datasets[1].data.push(ch2Value);

      // Manter apenas os últimos 20 pontos
      if (chartData.labels.length > 20) {
        chartData.labels.shift();
        chartData.datasets[0].data.shift();
        chartData.datasets[1].data.shift();
      }

      voltageChart.update('none');
    }

    function updateTelemetry() {
      const rollEl = document.getElementById("roll");
      const pitchEl = document.getElementById("pitch");
      const yawEl = document.getElementById("yaw");
      const cpCh1El = document.getElementById("cp_ch1");
      const cpCh2El = document.getElementById("cp_ch2");
      const depthEl = document.getElementById("depth");
      const cygnusEl = document.getElementById("cygnus");

      if (rollEl) {
        rollEl.innerHTML = getRandomAttitude() + '<span class="param-unit">°</span>';
        pitchEl.innerHTML = getRandomAttitude() + '<span class="param-unit">°</span>';
        yawEl.innerHTML = (Math.random() * 360).toFixed(1) + '<span class="param-unit">°</span>';
        cpCh1El.innerHTML = getRandomVoltage() + '<span class="param-unit">V</span>';
        cpCh2El.innerHTML = getRandomVoltage() + '<span class="param-unit">V</span>';
        depthEl.innerHTML = getRandomDepth() + '<span class="param-unit">m</span>';
        cygnusEl.innerHTML = getRandomCygnus() + '<span class="param-unit">%</span>';
      }

      // Atualizar gráfico se estiver visível
      if (voltageChart) {
        updateChartData();
      }

      // Atualizar valor do Cygnus na aba específica
      const cygnusReadingEl = document.getElementById("cygnus_reading");
      if (cygnusReadingEl) {
        cygnusReadingEl.innerHTML = getRandomReading() + '<span class="value-card-unit">mm</span>';
      }
    }

    // Atualizar a cada 500ms para parecer mais natural
    setInterval(updateTelemetry, 100);

    // Seções simplificadas
    const sections = {
      geral: `
        <div class="telemetry-table-container fade-in">
          <table class="telemetry-table">
            <thead>
              <tr>
                <th>Parâmetro</th>
                <th>Valor</th>
                <th>Status</th>
              </tr>
            </thead>
            <tbody>
              <tr>
                <td class="param-name">Roll</td>
                <td class="param-value" id="roll">0<span class="param-unit">°</span></td>
                <td class="param-status status-good">Ok</td>
              </tr>
              <tr>
                <td class="param-name">Pitch</td>
                <td class="param-value" id="pitch">0<span class="param-unit">°</span></td>
                <td class="param-status status-good">Ok</td>
              </tr>
              <tr>
                <td class="param-name">Yaw</td>
                <td class="param-value" id="yaw">0<span class="param-unit">°</span></td>
                <td class="param-status status-good">Ok</td>
              </tr>
              <tr>
                <td class="param-name">Deapth</td>
                <td class="param-value" id="depth">0<span class="param-unit">m</span></td>
                <td class="param-status status-good">Ok</td>
              </tr>              
              <tr>
                <td class="param-name">CP CH1</td>
                <td class="param-value" id="cp_ch1">0<span class="param-unit">V</span></td>
                <td class="param-status status-good">Ok</td>
              </tr>
              <tr>
                <td class="param-name">CP CH2</td>
                <td class="param-value" id="cp_ch2">0<span class="param-unit">V</span></td>
                <td class="param-status status-good">Ok</td>
              </tr>
              <tr>
                <td class="param-name">Cygnus</td>
                <td class="param-value" id="cygnus">0<span class="param-unit">%</span></td>
                <td class="param-status status-good">Ok</td>
              </tr>
            </tbody>
          </table>
        </div>
      `,
      cp: `
        <div class="chart-container fade-in">
          <canvas id="voltageChart"></canvas>
        </div>
        
        <div class="config-card fade-in">
          <div class="config-title">Configuração da Porta Serial</div>
          <div class="config-row">
            <div class="config-group">
              <label class="config-label">Porta COM</label>
              <select class="config-select" id="comPort">
                <option value="">Selecionar porta</option>
                <option value="COM1">COM1</option>
                <option value="COM2">COM2</option>
                <option value="COM3">COM3</option>
                <option value="COM4">COM4</option>
                <option value="COM5">COM5</option>
                <option value="COM6">COM6</option>
                <option value="COM7">COM7</option>
                <option value="COM8">COM8</option>
              </select>
            </div>
            
            <div class="config-group">
              <label class="config-label">Comunicação</label>
              <div class="checkbox-container">
                <input type="checkbox" class="config-checkbox" id="serialEnable">
                <span class="config-label">Habilitar comunicação</span>
              </div>
            </div>
            
            <button class="config-button" onclick="saveSerialConfig()">Salvar</button>
          </div>
        </div>
      `,
      cygnus: `
        <div class="section-title fade-in">Sonar Cygnus</div>
        <div class="value-card fade-in">
          <div class="value-card-label">Leitura</div>
          <div class="value-card-value" id="cygnus_reading">0<span class="value-card-unit">mm</span></div>
          <div class="value-card-description">Sensor Ultrassônico</div>
        </div>
        
        <div class="config-card fade-in">
          <div class="config-title">Configuração do Material</div>
          <div class="config-row">
            <div class="config-group">
              <label class="config-label">Lista de Materiais</label>
              <select class="config-select" id="materialSelect">
                <option value="">Selecionar material</option>
                <option value="aluminium-alloyed">Aluminium (alloyed)</option>
                <option value="aluminium-2014">Aluminium (2014)</option>
                <option value="aluminium-2024-t4">Aluminium (2024 T4)</option>
                <option value="aluminium-2117-t4">Aluminium (2117 T4)</option>
                <option value="brass-cuzn40">Brass (CuZn40)</option>
                <option value="brass-naval">Brass (Naval)</option>
                <option value="brass-cuzn30">Brass (CuZn30)</option>
                <option value="copper">Copper</option>
                <option value="grey-cast-iron">Grey Cast Iron</option>
                <option value="iconel">Iconel</option>
                <option value="lead">Lead</option>
                <option value="monel">Monel</option>
                <option value="nickel">Nickel</option>
                <option value="phosphor-bronze">Phosphor Bronze</option>
                <option value="mild-steel">Mild Steel</option>
                <option value="tool-steel">Tool Steel</option>
                <option value="stainless-steel-302">Stainless Steel 302</option>
                <option value="stainless-steel-347">Stainless Steel 347</option>
                <option value="stainless-steel-314">Stainless Steel 314</option>
                <option value="stainless-steel-316">Stainless Steel 316</option>
                <option value="f51-duplex-steel">F51 Duplex Steel UNS S31803</option>
                <option value="core-ten-steel">Core Ten Steel EN12223 S355-J0</option>
                <option value="tin">Tin</option>
                <option value="titanium">Titanium</option>
                <option value="tungsten-carbide">Tungsten Carbide</option>
                <option value="epoxy-resin">Epoxy Resin</option>
                <option value="acrylic">Acrylic</option>
                <option value="nylon-polyamide">Nylon (Polyamide)</option>
              </select>
            </div>
            
            <div class="config-group">
              <label class="config-label">Fator de Correção</label>
              <div class="input-with-unit">
                <input type="number" class="config-input" id="correctionFactor" placeholder="1.00" step="0.01" min="-200.0" max="200.0">
                <span class="unit-label">%</span>
              </div>
            </div>
            
            <button class="config-button" onclick="saveMaterialConfig()">Salvar</button>
          </div>
        </div>
      `
      //, joystick: `
      //   <div class="section-title fade-in">Controle Manual</div>
        
      //   <div class="joystick-cards-container">
        
      //     <div class="flight-mode-card fade-in">
      //       <div class="flight-mode-grid">
      //         <div class="flight-mode-button active" onclick="selectFlightMode('manual', this)">Manual</div>
      //         <div class="flight-mode-button" onclick="selectFlightMode('stabilize', this)">Stabilize</div>
      //         <div class="flight-mode-button" onclick="selectFlightMode('depth', this)">Depth</div>
      //         <div class="flight-mode-button" onclick="selectFlightMode('altitude', this)">Altitude</div>
      //       </div>
      //     </div>
          
      //     <div class="auxiliary-buttons-card fade-in">
      //       <div class="auxiliary-buttons-layout">
      //         <div class="auxiliary-top-button">
      //           <div class="auxiliary-button" onclick="selectAuxiliary('btn1', this)">Reset</div>
      //         </div>
      //         <div class="auxiliary-grid">
      //           <div class="auxiliary-button" onclick="selectAuxiliary('btn2', this)">Gain +</div>
      //           <div class="auxiliary-button" onclick="selectAuxiliary('btn3', this)">Jaw op.</div>
      //           <div class="auxiliary-button" onclick="selectAuxiliary('btn4', this)">Aux 1.</div>
      //           <div class="auxiliary-button" onclick="selectAuxiliary('btn5', this)">Gain -</div>
      //           <div class="auxiliary-button" onclick="selectAuxiliary('btn6', this)">Jaw cls.</div>
      //           <div class="auxiliary-button" onclick="selectAuxiliary('btn7', this)">Aux 2</div>
      //         </div>
      //       </div>
      //     </div>

      //     <div class="lights-card fade-in">
      //       <div class="lights-grid">
      //         <div class="lights-button" onclick="selectLight('light1+', this)">Light 1 +</div>
      //         <div class="lights-button" onclick="selectLight('light2+', this)">Light 2 +</div>
      //         <div class="lights-button" onclick="selectLight('light1-', this)">Light 1 -</div>
      //         <div class="lights-button" onclick="selectLight('light2-', this)">Light 2 -</div>
      //       </div>
      //     </div>
      //   </div>

      //   <!-- Trackbar simples -->
      //   <div class="slider-container fade-in">
      //     <input type="range" min="1" max="100" value="50" class="slider" id="myRange">
      //   </div>
      // `
    };

    function loadSection(section, element) {
      // Efeito de fade out
      const content = document.getElementById("content");
      content.style.opacity = '0';
      
      setTimeout(() => {
        content.innerHTML = sections[section];
        content.style.opacity = '1';
        
        // Atualizar menu ativo
        document.querySelectorAll('.menu-item').forEach(item => {
          item.classList.remove('active');
        });
        element.classList.add('active');

        // Inicializar gráfico se estivermos na seção CP
        if (section === 'cp') {
          setTimeout(() => {
            initVoltageChart();
          }, 100);
        }
      }, 150);
    }

    // Função para selecionar botões auxiliares
    function selectAuxiliary(button, element) {
      // Log da seleção (aqui você pode adicionar lógica específica)
      console.log('Botão auxiliar pressionado:', button);
      
      // Feedback visual
      element.style.transform = 'scale(0.9)';
      element.classList.add('active');
      
      setTimeout(() => {
        element.style.transform = '';
        element.classList.remove('active');
      }, 200);
    }

    // Função para controle das luzes
    function selectLight(light, element) {
      // Log da seleção
      console.log('Controle de luz:', light);
      
      // Feedback visual momentâneo
      element.style.transform = 'scale(0.9)';
      element.classList.add('active');
      
      setTimeout(() => {
        element.style.transform = '';
        element.classList.remove('active');
      }, 200);
    }

    // Função para salvar configuração serial
    function saveSerialConfig() {
      const comPort = document.getElementById('comPort');
      const serialEnable = document.getElementById('serialEnable');
      
      const selectedPort = comPort.value;
      const isEnabled = serialEnable.checked;
      
      if (!selectedPort) {
        alert('Por favor, selecione uma porta COM.');
        return;
      }
      
      // Aqui você pode adicionar a lógica para salvar as configurações
      console.log('Configuração salva:', {
        porta: selectedPort,
        habilitado: isEnabled
      });
      
      // Feedback visual (opcional)
      const button = document.querySelector('.config-button');
      const originalText = button.textContent;
      button.textContent = 'Salvo!';
      button.style.background = 'linear-gradient(135deg, #22c55e, #16a34a)';
      
      setTimeout(() => {
        button.textContent = originalText;
        button.style.background = 'linear-gradient(135deg, #3b82f6, #2563eb)';
      }, 2000);
    }

    // Função para selecionar Flight Mode
    function selectFlightMode(mode, element) {
      // Remover classe active de todos os botões
      document.querySelectorAll('.flight-mode-button').forEach(btn => {
        btn.classList.remove('active');
      });
      
      // Adicionar classe active ao botão clicado
      element.classList.add('active');
      
      // Log da seleção (aqui você pode adicionar lógica específica)
      console.log('Flight Mode selecionado:', mode);
      
      // Feedback visual adicional (opcional)
      element.style.transform = 'scale(0.95)';
      setTimeout(() => {
        element.style.transform = '';
      }, 150);
    }

    // Função para salvar configuração de material
    function saveMaterialConfig() {
      const materialSelect = document.getElementById('materialSelect');
      const correctionFactor = document.getElementById('correctionFactor');
      
      const selectedMaterial = materialSelect.value;
      const selectedText = materialSelect.options[materialSelect.selectedIndex].text;
      const factorValue = correctionFactor.value || '1.00';
      
      if (!selectedMaterial) {
        alert('Por favor, selecione um material.');
        return;
      }
      
      // Validar fator de correção
      const factor = parseFloat(factorValue);
      if (isNaN(factor) || factor < -200.0 || factor > 200.0) {
        alert('Fator de correção deve estar entre -200.0% e 200.0%');
        return;
      }
      
      // Aqui você pode adicionar a lógica para salvar a configuração do material
      console.log('Configuração salva:', {
        material: {
          valor: selectedMaterial,
          nome: selectedText
        },
        fatorCorrecao: factor
      });
      
      // Feedback visual
      const button = document.querySelector('.config-card .config-button');
      const originalText = button.textContent;
      button.textContent = 'Salvo!';
      button.style.background = 'linear-gradient(135deg, #22c55e, #16a34a)';
      
      setTimeout(() => {
        button.textContent = originalText;
        button.style.background = 'linear-gradient(135deg, #3b82f6, #2563eb)';
      }, 2000);
    }

    // Inicialização
    document.addEventListener('DOMContentLoaded', function() {
      updateTelemetry();
    });