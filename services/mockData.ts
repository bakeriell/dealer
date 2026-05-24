import { DealerData, ColumnDefinition } from '../types';

export const generateMockData = (count: number = 30): { data: DealerData[], columns: ColumnDefinition[], assignedMap: Record<string, number> } => {
  const dealers = ['DEALER A', 'DEALER B', 'DEALER C', 'DEALER D', 'DEALER E'];
  const models = ['ATTO 3', 'SEAL', 'DOLPHIN', 'TANG', 'HAN'];
  const competitors = ['Tesla Model 3', 'MG 4', 'Volkswagen ID.3', 'Nessuno', 'Polestar 2', 'Volvo EX30', 'Tesla Model Y', 'No', 'Nessuno'];
  const sources = ['Facebook', 'Google Ads', 'Organic Website', 'Event', 'Newsletter', 'Instagram'];

  // Specific answers for the purchase question
  const purchaseAnswers = [
    'Si, BYD',
    'No, sono ancora indeciso',
    'No, ho deciso di rimandare per ora l\'acquisto',
    'Si, ho acquistato un altro Brand/Modello'
  ];

  // Specific answers for Contact Question (to match screenshot)
  const contactAnswers = [
    'Sì', 'Sì', 'Sì', 'Sì', // Weighted towards Yes
    'No', 
    'No, mi sono recato in autonomia in concessionaria',
    'Non ricordo di aver fatto questa richiesta'
  ];

  const responseAnswers = [
    'Sì', 'Sì', 'Sì', 
    'No', 
    'Sì, ma non vuole rispondere al sondaggio'
  ];

  const data: DealerData[] = Array.from({ length: count }).map((_, index) => {
    const dealer = dealers[Math.floor(Math.random() * dealers.length)];
    const model = models[Math.floor(Math.random() * models.length)];
    const source = sources[Math.floor(Math.random() * sources.length)];
    
    // Q1: User Responded? (The funnel entry point)
    const respondedStatus = responseAnswers[Math.floor(Math.random() * responseAnswers.length)];
    const isSurveyTaken = respondedStatus === 'Sì';

    // Q2: Remember Request?
    const remember = isSurveyTaken ? (Math.random() > 0.1 ? 'Sì' : 'No') : '';

    // Q3: Contacted? (The detailed question)
    const contactedRaw = isSurveyTaken ? contactAnswers[Math.floor(Math.random() * contactAnswers.length)] : '';
    const isContactedPositive = contactedRaw === 'Sì';
    
    const visited = isContactedPositive && Math.random() > 0.5 ? 'Sì' : 'No';

    // Quote Logic (Dependent on contact)
    const quoteGiven = isContactedPositive && Math.random() > 0.2 ? 'Sì' : 'No';
    const price = quoteGiven === 'Sì' ? Math.floor(30000 + Math.random() * 25000) : null;
    const consistency = quoteGiven === 'Sì' ? (Math.random() > 0.2 ? 'Sì' : 'No') : '';

    // Finalized Purchase Logic
    let finalized = '';
    if (quoteGiven === 'Sì') {
      const rand = Math.random();
      if (rand > 0.7) finalized = purchaseAnswers[0]; // Si, BYD
      else if (rand > 0.5) finalized = purchaseAnswers[1]; // Indeciso
      else if (rand > 0.3) finalized = purchaseAnswers[2]; // Rimandato
      else finalized = purchaseAnswers[3]; // Altro Brand
    } else {
      finalized = isContactedPositive ? purchaseAnswers[1] : ''; 
    }
    
    // Competitors
    const compCount = Math.floor(Math.random() * 2);
    const selectedComps = [];
    if (Math.random() > 0.3) {
        for(let i=0; i<=compCount; i++) {
            selectedComps.push(competitors[Math.floor(Math.random() * competitors.length)]);
        }
    } else {
        selectedComps.push('Nessuno');
    }
    const otherBrands = [...new Set(selectedComps)].join(', ');

    return {
      id: index.toString(),
      'Lead Number': `LD${String(index + 1).padStart(3, '0')}`,
      'Dealer Name': dealer,
      'Model': model,
      'Fonte': source, // Added Source
      'L’utente ha risposto': respondedStatus, // New Question
      'Ha effettuato una richiesta per un test drive?': remember,
      'Dopo aver lasciato i suoi dati, è stato contattato da un concessionario BYD?': contactedRaw,
      'Visited Dealership': visited,
      'Il concessionario le ha fornito un preventivo per il modello di suo interesse?': quoteGiven,
      'Si ricorda quale prezzo le è stato proposto dal concessionario?': price ? String(price) : '',
      'Ha finalizzato l’acquisto dell\'auto di suo interesse?': finalized,
      'Sta valutando altri Brand/Modelli?': otherBrands,
      'Il preventivo che le è stato offerto dal concessionario è coerente con quanto comunicato da BYD?': consistency,
      'Date': new Date(2024, 0, Math.floor(Math.random() * 28) + 1).toISOString().split('T')[0]
    };
  });

  const firstRecord = data[0];
  const columns: ColumnDefinition[] = Object.keys(firstRecord).map(key => ({
    key,
    label: key,
    type: typeof firstRecord[key] === 'number' ? 'number' : 'text'
  }));

  // Mock Assigned Leads Map (Simulating Sheet 2)
  const assignedMap: Record<string, number> = {};
  dealers.forEach(d => {
    assignedMap[d] = Math.floor(count / dealers.length) + Math.floor(Math.random() * 5); 
  });

  return { data, columns, assignedMap };
};