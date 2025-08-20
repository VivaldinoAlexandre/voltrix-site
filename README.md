[index.html](https://github.com/user-attachments/files/21903338/index.html)
<script type="text/javascript">
        var gk_isXlsx = false;
        var gk_xlsxFileLookup = {};
        var gk_fileData = {};
        function filledCell(cell) {
          return cell !== '' && cell != null;
        }
        function loadFileData(filename) {
        if (gk_isXlsx && gk_xlsxFileLookup[filename]) {
            try {
                var workbook = XLSX.read(gk_fileData[filename], { type: 'base64' });
                var firstSheetName = workbook.SheetNames[0];
                var worksheet = workbook.Sheets[firstSheetName];

                // Convert sheet to JSON to filter blank rows
                var jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, blankrows: false, defval: '' });
                // Filter out blank rows (rows where all cells are empty, null, or undefined)
                var filteredData = jsonData.filter(row => row.some(filledCell));

                // Heuristic to find the header row by ignoring rows with fewer filled cells than the next row
                var headerRowIndex = filteredData.findIndex((row, index) =>
                  row.filter(filledCell).length >= filteredData[index + 1]?.filter(filledCell).length
                );
                // Fallback
                if (headerRowIndex === -1 || headerRowIndex > 25) {
                  headerRowIndex = 0;
                }
                    
                // Convert filtered JSON back to CSV
                var csv = XLSX.utils.aoa_to_sheet(filteredData.slice(headerRowIndex)); // Create a new sheet from filtered array of arrays
                csv = XLSX.utils.sheet_to_csv(csv, { header: 1 });
                return csv;
            } catch (e) {
                console.error(e);
                return "";
            }
        }
        return gk_fileData[filename] || "";
        }
        </script><!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Voltrix - Instalações Elétricas</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        body {
            font-family: 'Arial', sans-serif;
            background-color: #000000;
            color: #ffffff;
            line-height: 1.8;
            font-size: 1.2em;
        }
        header {
            background-color: #ff6200;
            color: #000000;
            padding: 1.5em;
            text-align: center;
            position: relative;
        }
        header img.logo {
            max-width: 100%;
            height: 150px;
            max-height: 150px; /* Ajuste o tamanho máximo da logo, se necessário */
        }
        header p {
            font-size: 2em;
            margin-top: 0.5em;
        }
        .date {
            font-size: 0.9em;
            margin-top: 0.5em;
            color: #000000;
        }
        nav {
            background-color: #ff6200;
            padding: 1em;
            text-align: center;
        }
        nav a {
            color: #000000;
            text-decoration: none;
            margin: 0 1.5em;
            font-weight: bold;
            font-size: 1.3em;
            transition: color 0.3s;
        }
        nav a:hover {
            color: #ffffff;
        }
        main {
            max-width: 1000px;
            margin: 0 auto;
            padding: 2em;
        }
        section {
            margin-bottom: 3em;
        }
        h2 {
            color: #ff6200;
            margin-bottom: 1em;
            font-size: 2.5em;
        }
        .services {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 1.5em;
        }
        .service-card {
            background-color: #1a1a1a;
            padding: 1.5em;
            border-radius: 8px;
            box-shadow: 0 2px 5px rgba(255, 98, 0, 0.2);
            text-align: center;
        }
        .service-card h3 {
            color: #ff6200;
            margin-bottom: 0.5em;
            font-size: 1.8em;
        }
        .contact-form {
            max-width: 500px;
            margin: 0 auto;
            display: flex;
            flex-direction: column;
            gap: 1em;
        }
        .contact-form input, .contact-form textarea {
            padding: 1em;
            border: 2px solid #ff6200;
            border-radius: 8px;
            font-size: 1.2em;
            background-color: #333333;
            color: #ffffff;
        }
        .contact-form button {
            padding: 1em;
            background-color: #ff6200;
            color: #000000;
            border: none;
            border-radius: 8px;
            cursor: pointer;
            font-size: 1.3em;
            transition: background-color 0.3s;
        }
        .contact-form button:hover {
            background-color: #ffffff;
            color: #ff6200;
        }
        .whatsapp, .assistant {
            margin-top: 1.5em;
            text-align: center;
        }
        .whatsapp a, .assistant a {
            color: #ff6200;
            text-decoration: none;
            font-weight: bold;
            font-size: 1.3em;
        }
        .whatsapp a:hover, .assistant a:hover {
            text-decoration: underline;
        }
        .assistant {
            background-color: #1a1a1a;
            padding: 1.2em;
            border-radius: 10px;
            box-shadow: 0 2px 5px rgba(255, 98, 0, 0.2);
            margin: 1.5em auto;
            max-width: 350px;
        }
        .assistant a {
            display: inline-block;
            padding: 0.7em 1.2em;
            background-color: #ff6200;
            color: #000000;
            border-radius: 8px;
        }
        .assistant a:hover {
            background-color: #ffffff;
        }
        footer {
            background-color: #ff6200;
            color: #000000;
            text-align: center;
            padding: 1em;
            width: 100%;
            font-size: 1.2em;
        }
        @media (max-width: 600px) {
            header img.logo {
                max-width: 90%;
            }
            nav a {
                display: block;
                margin: 0.8em 0;
                font-size: 1.2em;
            }
            h2 {
                font-size: 2em;
            }
            .service-card h3 {
                font-size: 1.5em;
            }
            header p {
                font-size: 1.5em;
            }
        }
    </style>
</head>
<body>
    <header>
        <img src="logo 1.png" alt="Logo Voltrix" class="logo">
        <p>Soluções completas em elétrica industrial e predial</p>
        <div class="date">Atualizado em: Quarta-feira, 20 de Agosto de 2025, 15:34 PM -03</div>
    </header>
    <nav>
        <a href="#inicio">Início</a>
        <a href="#servicos">Serviços</a>
        <a href="#contato">Contato</a>
    </nav>
    <main>
        <section id="inicio">
            <h2>Bem-vindo à Voltrix</h2>
            <p>Oferecemos serviços de instalação elétrica com qualidade e segurança, atendendo às normas ABNT NBR 5410. Especialistas em projetos industriais e prediais.</p>
        </section>
        <section id="servicos">
            <h2>Nossos Serviços</h2>
            <div class="services">
                <div class="service-card">
                    <h3>Montagem de Quadros Elétricos</h3>
                    <p>Instalação de quadros de distribuição com alta confiabilidade.</p>
                </div>
                <div class="service-card">
                    <h3>Instalações Elétricas</h3>
                    <p>Soluções industriais e prediais sob medida.</p>
                </div>
                <div class="service-card">
                    <h3>Montagem de Caixas de Tomadas</h3>
                    <p>Projetos personalizados de alta qualidade.</p>
                </div>
                <div class="service-card">
                    <h3>Manutenção em Quadros</h3>
                    <p>Inspeções e reparos para máxima eficiência.</p>
                </div>
            </div>
        </section>
        <section id="contato">
            <h2>Entre em Contato</h2>
            <p>Precisa de uma solução elétrica? Fale com nosso assistente ou envie sua mensagem!</p>
            <form class="contact-form">
                <input type="text" placeholder="Seu Nome" required>
                <input type="email" placeholder="Seu E-mail" required>
                <input type="tel" placeholder="Seu Telefone" required>
                <textarea placeholder="Sua Mensagem" rows="5" required></textarea>
                <button type="button" onclick="enviarFormulario()">Enviar Mensagem</button>
            </form>
            <div class="whatsapp">
                <p>Ou entre em contato via WhatsApp: <a href="https://wa.me/5599696224416" target="_blank">(67) 99622-4416</a></p>
            </div>
            <div class="assistant">
                <p>Fale com nosso assistente virtual: <a href="https://wa.me/5599696224416?text=Ol%C3%A1,%20gostaria%20de%20mais%20informa%C3%A7%C3%B5es%20sobre%20seus%20servi%C3%A7os!" target="_blank">Chamar Assistente</a></p>
            </div>
        </section>
    </main>
    <footer>
        <p>&copy; 2025 Voltrix - Instalações Elétricas. Todos os direitos reservados.</p>
    </footer>
    <script>
        function enviarFormulario() {
            alert("Mensagem enviada com sucesso! Entraremos em contato em breve. (Este é um exemplo, nenhuma mensagem foi realmente enviada.)");
        }
    </script>
</body>
</html>
