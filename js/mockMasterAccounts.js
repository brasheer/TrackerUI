// ==========================================================================
// TrackerUI — Mock Master Accounts Dataset (Testing & Localhost Dev Mode)
// ==========================================================================
// Used when running in localhost / Dev Mock Mode to test bidirectional
// Broker <-> Master # auto-fill, standalone Yard # / Name dropdowns, and UUIDv4
// request workflows without requiring a real 'Master Accounts' Excel sheet.

window.MOCK_MASTER_ACCOUNTS = Object.freeze([
  {
    broker: 'Auto Import Tbilisi LLC',
    master: 'MA-104820'
  },
  {
    broker: 'Caucasus Transit Motors LLC',
    master: 'MA-109340'
  },
  {
    broker: 'Baltic Sea Remarketing SIA',
    master: 'MA-203915'
  },
  {
    broker: 'Atlantic Port Export Group',
    master: 'MA-318402'
  },
  {
    broker: 'Sunbelt Fleet Liquidators Inc',
    master: 'MA-405128'
  },
  {
    broker: 'Great Lakes Auto Exchange LLC',
    master: 'MA-512790'
  },
  {
    broker: 'Lone Star Salvage & Export Co',
    master: 'MA-629044'
  },
  {
    broker: 'Pacific Rim Vehicle Logistics',
    master: 'MA-734119'
  },
  {
    broker: 'Tri-State Dealer Services LLC',
    master: 'MA-840662'
  },
  {
    broker: 'Meridian Global Auto Traders',
    master: 'MA-951208'
  }
]);

window.MOCK_YARD_LIST = Object.freeze([
  { value: '108', name: 'Atlanta West' },
  { value: '110', name: 'Atlanta East' },
  { value: '145', name: 'Savannah / Port' },
  { value: '204', name: 'Newark / NY Metro' },
  { value: '225', name: 'Houston North' },
  { value: '312', name: 'Chicago South' },
  { value: '340', name: 'Dallas / Fort Worth' },
  { value: '418', name: 'Los Angeles / Long Beach' }
]);

