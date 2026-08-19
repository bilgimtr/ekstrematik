import { getStore } from '@netlify/blobs';

function todayKey(){
  return new Date().toISOString().slice(0, 10); // YYYY-MM-DD (UTC)
}

function parseCookie(header, name){
  if(!header) return null;
  const match = header.match(new RegExp('(?:^|; )' + name + '=([^;]*)'));
  return match ? decodeURIComponent(match[1]) : null;
}

function makeId(){
  return crypto.randomUUID();
}

export default async (req) => {
  const visitors = getStore('visitors', { consistency: 'strong' });
  const stats = getStore('stats', { consistency: 'strong' });

  const cookieHeader = req.headers.get('cookie');
  let visitorId = parseCookie(cookieHeader, 'ekv_id');
  const today = todayKey();

  let isNewVisitor = false;
  let isNewToday = false;

  if(!visitorId){
    visitorId = makeId();
    isNewVisitor = true;
    isNewToday = true;
  } else {
    const record = await visitors.get(visitorId, { type: 'json' });
    if(!record){
      isNewVisitor = true;
      isNewToday = true;
    } else if(record.lastSeen !== today){
      isNewToday = true;
    }
  }

  if(isNewVisitor || isNewToday){
    await visitors.setJSON(visitorId, { lastSeen: today });
  }

  if(isNewVisitor){
    const totalRaw = await stats.get('total');
    const total = (parseInt(totalRaw, 10) || 0) + 1;
    await stats.set('total', String(total));
  }
  if(isNewToday){
    const dayKey = 'day:' + today;
    const dayRaw = await stats.get(dayKey);
    const dayCount = (parseInt(dayRaw, 10) || 0) + 1;
    await stats.set(dayKey, String(dayCount));
  }

  const totalNow = parseInt(await stats.get('total'), 10) || 0;
  const todayNow = parseInt(await stats.get('day:' + today), 10) || 0;

  const headers = {
    'Content-Type': 'application/json',
    'Cache-Control': 'no-store',
  };
  if(isNewVisitor){
    // Kimliksiz, rastgele bir sayaç kimliği — kişisel veri değil, sadece
    // aynı ziyaretçinin bugün/toplamda birden fazla kez sayılmasını önler.
    headers['Set-Cookie'] = `ekv_id=${visitorId}; Max-Age=63072000; Path=/; SameSite=Lax; Secure`;
  }

  return new Response(JSON.stringify({ total: totalNow, today: todayNow }), {
    status: 200,
    headers,
  });
};

export const config = { path: '/api/visit' };
