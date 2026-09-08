// =========================================================
// ELEMENTOS DA INTERFACE
// =========================================================

const reservaButton =
    document.getElementById("botao-reserva");

const statusText =
    document.getElementById("statusText");

const statusIndicator =
    document.getElementById("statusIndicator");

const log =
    document.getElementById("log");

const clearLog =
    document.getElementById("clearLog");


// =========================================================
// ADICIONAR LOG
// =========================================================

function adicionarLog(mensagem, tipo = "info") {

    if (!log) {
        console.log(`[${tipo}] ${mensagem}`);
        return;
    }

    const linha =
        document.createElement("div");

    linha.classList.add("log-line");

    if (tipo) {
        linha.classList.add(tipo);
    }

    const horario =
        new Date().toLocaleTimeString();

    linha.textContent =
        `[${horario}] ${mensagem}`;

    log.appendChild(linha);

    log.scrollTop =
        log.scrollHeight;
}


// =========================================================
// ALTERAR STATUS
// =========================================================

function alterarStatus(texto, cor) {

    if (statusText) {
        statusText.textContent = texto;
    }

    if (statusIndicator) {

        statusIndicator.style.backgroundColor =
            cor;

        statusIndicator.style.boxShadow =
            `0 0 10px ${cor}`;
    }
}


// =========================================================
// EXECUTAR RESERVA
// =========================================================

async function executarReserva() {

    if (!reservaButton) {
        return;
    }


    // =====================================================
    // DESABILITAR BOTÃO
    // =====================================================

    reservaButton.disabled = true;

    reservaButton.innerHTML = `
        <span class="button-icon">⏳</span>
        EXECUTANDO...
    `;


    // =====================================================
    // STATUS
    // =====================================================

    alterarStatus(
        "Executando reservas...",
        "#f59e0b"
    );


    // =====================================================
    // LOG
    // =====================================================

    adicionarLog(
        "Iniciando processo de reservas...",
        "info"
    );


    try {

        // =================================================
        // VERIFICAR PYWEBVIEW
        // =================================================

        if (
            typeof pywebview === "undefined" ||
            !pywebview.api
        ) {

            throw new Error(
                "PyWebView não está disponível."
            );
        }


        // =================================================
        // CHAMAR PYTHON
        // =================================================

        const resultado =
            await pywebview.api.executar_reserva();


        console.log(
            "Resposta do Python:",
            resultado
        );


        // =================================================
        // VERIFICAR RESULTADO
        // =================================================

        if (resultado && resultado.sucesso) {

            adicionarLog(
                resultado.mensagem,
                "success"
            );

            alterarStatus(
                "Processo iniciado",
                "#00c853"
            );


            reservaButton.innerHTML = `
                <span class="button-icon">✓</span>
                RESERVA INICIADA
            `;

        } else {

            const mensagem =
                resultado?.mensagem ||
                "Ocorreu um erro ao iniciar a reserva.";

            adicionarLog(
                mensagem,
                "error"
            );

            alterarStatus(
                "Erro",
                "#ff4d4d"
            );


            reservaButton.innerHTML = `
                <span class="button-icon">✕</span>
                ERRO
            `;
        }


    } catch (erro) {

        // =================================================
        // ERRO
        // =================================================

        console.error(
            "Erro:",
            erro
        );


        adicionarLog(
            "Erro ao comunicar com o Python.",
            "error"
        );


        adicionarLog(
            String(erro),
            "error"
        );


        alterarStatus(
            "Erro",
            "#ff4d4d"
        );


        reservaButton.innerHTML = `
            <span class="button-icon">✕</span>
            ERRO
        `;
    }


    // =====================================================
    // RESTAURAR BOTÃO
    // =====================================================

    setTimeout(
        function () {

            reservaButton.disabled =
                false;

            reservaButton.innerHTML = `
                <span class="button-icon">▶️</span>
                RESERVA EXPORTAR
            `;

        },
        2000
    );
}


// =========================================================
// BOTÃO — RESERVA EXPORTAR
// =========================================================

if (reservaButton) {

    reservaButton.addEventListener(
        "click",
        executarReserva
    );

} else {

    console.error(
        "Botão #botao-reserva não encontrado no HTML."
    );
}


// =========================================================
// LIMPAR LOG
// =========================================================

if (clearLog) {

    clearLog.addEventListener(
        "click",
        function () {

            if (log) {

                log.innerHTML = "";

                adicionarLog(
                    "Log limpo.",
                    "info"
                );
            }

        }
    );
}


// =========================================================
// MENSAGEM INICIAL
// =========================================================

adicionarLog(
    "Sistema VIX pronto.",
    "success"
);