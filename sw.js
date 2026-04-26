// ══════════════════════════════════════════════════════
//  INSTITUTO BÍBLICO ZAO — Service Worker
//  Para atualizar o cache em todos os dispositivos,
//  basta incrementar o número da versão abaixo.
// ══════════════════════════════════════════════════════

var VERSAO = 'zao-v4';

var ARQUIVOS = [
  './',
  './ZAO_Presenca.html',
  'https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap',
  'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'
];

// Instalação — salva os arquivos no cache
self.addEventListener('install', function(event) {
  event.waitUntil(
    caches.open(VERSAO).then(function(cache) {
      return cache.addAll(ARQUIVOS);
    }).then(function() {
      // Ativa imediatamente sem esperar a aba fechar
      return self.skipWaiting();
    })
  );
});

// Ativação — apaga caches de versões antigas
self.addEventListener('activate', function(event) {
  event.waitUntil(
    caches.keys().then(function(chaves) {
      return Promise.all(
        chaves.filter(function(chave) {
          return chave !== VERSAO;
        }).map(function(chave) {
          return caches.delete(chave);
        })
      );
    }).then(function() {
      // Assume controle de todas as abas abertas imediatamente
      return self.clients.claim();
    })
  );
});

// Intercepta requisições — serve do cache, atualiza em segundo plano
self.addEventListener('fetch', function(event) {
  // Ignora requisições que não sejam GET
  if (event.request.method !== 'GET') return;

  // Para a API do QR Code e Google Sheets — sempre vai para a rede
  var url = event.request.url;
  if (url.indexOf('qrserver.com') >= 0 || url.indexOf('script.google.com') >= 0) {
    event.respondWith(fetch(event.request));
    return;
  }

  // Estratégia "Cache primeiro, atualiza em segundo plano"
  event.respondWith(
    caches.open(VERSAO).then(function(cache) {
      return cache.match(event.request).then(function(resposta) {
        // Busca versão nova em segundo plano sempre
        var fetchPromise = fetch(event.request).then(function(respostaRede) {
          if (respostaRede && respostaRede.status === 200) {
            cache.put(event.request, respostaRede.clone());
          }
          return respostaRede;
        }).catch(function() {
          // Sem internet — usa o cache silenciosamente
        });

        // Retorna cache imediatamente se disponível, senão espera a rede
        return resposta || fetchPromise;
      });
    })
  );
});

// Recebe mensagem do app para forçar atualização imediata
self.addEventListener('message', function(event) {
  if (event.data === 'pular-espera') {
    self.skipWaiting();
  }
});
