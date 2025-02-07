from icalendar import Calendar

def tipo_calendario(content):
    try:
        # Decodificar o conteúdo do calendário
        ical = Calendar.from_ical(content)
        eventos = []

        for componente in ical.walk():
            if componente.name == "VEVENT":
                #Informações do evento
                titulo = componente.get("summary", "Sem Título")
                inicio = componente.get("dtstart").dt
                fim = componente.get("dtend").dt
                descricao = componente.get("description", "Sem descrição")
                local = componente.get("location", "Sem local")

                eventos.append({
                    "titulo": titulo,
                    "inicio": inicio,
                    "fim": fim,
                    "descricao": descricao,
                    "local": local
                })

        #HTML para os eventos
        html_eventos = "<h3>Eventos de Calendário:</h3><ul>"
        for evento in eventos:
            html_eventos += f"""
            <li>
                <strong>{evento['titulo']}</strong><br>
                Início: {evento['inicio']}<br>
                Fim: {evento['fim']}<br>
                Local: {evento['local']}<br>
                Descrição: {evento['descricao']}
            </li>
            """
        html_eventos += "</ul>"

        return html_eventos

    except Exception as e:
        print(f"Erro ao processar calendário: {e}")
        return "<p>Não foi possível processar o conteúdo do calendário.</p>"