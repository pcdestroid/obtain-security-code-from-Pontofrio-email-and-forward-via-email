function codigo_pontofrio() {
  const assuntoAProcurar = 'Código de segurança Pontofrio';

  // Procurar apenas uma thread com o assunto específico.
  const threads = GmailApp.search(`is:unread subject:"${assuntoAProcurar}"`, 0, 1);
  Logger.log(threads.length);

  if (threads.length > 0 && threads[0].isUnread()) { // Se houver pelo menos uma thread encontrada e não foi lido.
    const messages = threads[0].getMessages();

    // Obter o corpo do primeiro e-mail na thread
    const corpoEmail = messages[0].getPlainBody(); // Ou use messages[0].getBody() se preferir o corpo HTML

    // Usar expressão regular para encontrar o código
    const regexCodigo = /O seu código é (\d+)/;
    const match = corpoEmail.match(regexCodigo);

    threads[0].markRead();

    let emailComprador = CalendarApp.getId();
    let dadosComprador = getUserByEmail(emailComprador)
    let assinatura = `
  
  <table style="color: rgb(136, 136, 136); white-space: normal; background-color: rgb(255, 255, 255); font-size: 12px; font-family: Arial; width: 513px;">
  <tbody>
  <tr>
  <td style="margin: 0px; max-width: 100px; width: 100px;" valign="top"><img class="CToWUd" style="border-radius: 4px; width: 113px; height: 118px;" src="https://ci6.googleusercontent.com/proxy/iEOrmj7yb3xF4gNx9mDtAxGFuveeaACz1ZAt-RYPn78dDAGhfRJDLwhu34hjsDUlpuRooQ_aF2V0DcCRjx4IDp3O=s0-d-e1-ft#https://i.ibb.co/fG73FxC/Logo-BR-Assinatura.png" alt="Sua foto ou logo" data-bit="iit"></td>
  <td style="margin: 0px; padding: 0px 8px; width: 414px;">
  <table style="color: rgb(128, 128, 128); height: 108px; width: 414px;">
  <tbody>
  <tr style="height: 18px;">
  <td style="margin: 0px; font-size: 14px; font-weight: bold; color: rgb(14, 43, 141); height: 18px; width: 404px;">${dadosComprador[0]}</td>
  </tr>
  <tr style="height: 18px;">
  <td style="margin: 0px; height: 18px; width: 404px;">Compras</td>
  </tr>
  <tr style="height: 18px;">
  <td style="margin: 0px; height: 18px; width: 404px;">BR Marinas</td>
  </tr>
  <tr style="height: 18px;">
  <td style="margin: 0px; height: 18px; width: 404px;">${dadosComprador[5]}&nbsp;&nbsp;<a aria-label="Chat on WhatsApp" href="https://wa.me/${dadosComprador[6]}"><img class="CToWUd" src="https://ci3.googleusercontent.com/proxy/kJDPaPYcNQs64k_qKrGFp6XYuYXrA0FkVNTRpvuAzRv7COIpd65R8420WBcSTG4QRUPKdM_7DUQ=s0-d-e1-ft#https://i.ibb.co/Bjn8dcj/whatsapp.png" alt="Whatsapp" width="16" height="16" data-bit="iit"></a></td>
  </tr>
  <tr style="height: 18px;">
  <td style="margin: 0px; height: 18px; width: 404px;"><a style="color: rgb(17, 85, 204);" href="mailto:${emailComprador}" target="_blank" rel="noopener">${emailComprador}</a></td>
  </tr>
  </tbody>
  </table>
  </td>
  </tr>
  </tbody>
  </table>
  `



    const email_para_envio = PropertiesService.getScriptProperties().getProperty("email_para_envio");

    if (match && match[1]) {
      const codigo = match[1];
      Logger.log(`Código encontrado: ${codigo}`);

      let emailTemp = HtmlService.createHtmlOutput(`<span>${bomDia()}!</span><br><span> Segue código da Pontofrio: ${codigo}<br><br>Atenciosamente,</span><br><br>${assinatura}<br>`).setTitle('Email');
      let htmlMessage = emailTemp.getContent();
      // Crie o rascunho do email com anexo
      const subject = `Código Pontofrio - ${codigo} `;

      sendMessage(`Segue código da Ponto Frio: ${codigo}`)

      GmailApp.sendEmail(
        email_para_envio,
        subject,
        "Your email doesn't support HTML.",
        {
          name: 'Código',
          htmlBody: htmlMessage,
          bcc: ''
        }
      );

      // Agora você pode fazer o que quiser com o código, por exemplo, armazená-lo em uma variável ou enviá-lo para outro lugar.
    } else {
      Logger.log("Código não encontrado");
    }
  } else {
    Logger.log("Nada encontrado");
  }
}

function getUserByEmail(email) {
  //Pegar nome do usuário
  const id_plan_dados_usuarios = PropertiesService.getScriptProperties().getProperty("id_plan_dados_usuarios")
  let users = SpreadsheetApp.openById(id_plan_dados_usuarios).getSheetByName('usuarios');
  let dadosUsers = users.getRange(3, 1, users.getLastRow(), 7).getValues();
  //Pegar nome do comprador
  let user = dadosUsers.find(dado => dado[2] === email);
  return user ? user : '';
}

function sendMessage(text) {

  const keyWebHook = PropertiesService.getScriptProperties().getProperty("keyWebHook")
  const GOOGLE_CHAT_WEBHOOK_LINK = `https://chat.googleapis.com/v1/spaces/AAAAMJJMeOA/messages?key=${keyWebHook}`;
  const payload = JSON.stringify({ text: text });
  const options = {
    method: 'POST',
    contentType: 'application/json',
    payload: payload,
  };
  UrlFetchApp.fetch(GOOGLE_CHAT_WEBHOOK_LINK, options);
  return true;
}

function bomDia() {
  // Retorna a frase, bom dia quando é de dia, boa tarde quando é de tarde e boa noite quando é de noite.
  let data = new Date();
  let hora = data.getHours();
  switch (true) {
    case (hora >= 5 && hora < 12):
      return 'Bom dia';
    case (hora >= 12 && hora < 18):
      return 'Boa tarde';
    case (hora >= 18 && hora <= 24 || hora >= 0 && hora <= 4):
      return 'Boa noite';
  }
}
