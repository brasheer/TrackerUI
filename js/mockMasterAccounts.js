// ==========================================================================
// TrackerUI — Mock Master Accounts Dataset (Testing & Localhost Dev Mode)
// ==========================================================================
// Used when running in localhost / Dev Mock Mode to test bidirectional
// Broker <-> Master # auto-fill, Bidder # / Yard # defaults, and UUIDv4
// request workflows without requiring a real 'Master Accounts' Excel sheet.

window.MOCK_MASTER_ACCOUNTS = Object.freeze([
  {
    broker: 'Auto Import Tbilisi LLC',
    master: 'MA-104820',
    bidder: 'BD-882104',
    yard: '110'
  },
  {
    broker: 'Caucasus Transit Motors LLC',
    master: 'MA-109340',
    bidder: 'BD-884512',
    yard: '110'
  },
  {
    broker: 'Baltic Sea Remarketing SIA',
    master: 'MA-203915',
    bidder: 'BD-741029',
    yard: '145'
  },
  {
    broker: 'Atlantic Port Export Group',
    master: 'MA-318402',
    bidder: 'BD-619384',
    yard: '204'
  },
  {
    broker: 'Sunbelt Fleet Liquidators Inc',
    master: 'MA-405128',
    bidder: 'BD-530291',
    yard: '110'
  },
  {
    broker: 'Great Lakes Auto Exchange LLC',
    master: 'MA-512790',
    bidder: 'BD-492817',
    yard: '312'
  },
  {
    broker: 'Lone Star Salvage & Export Co',
    master: 'MA-629044',
    bidder: 'BD-384920',
    yard: '225'
  },
  {
    broker: 'Pacific Rim Vehicle Logistics',
    master: 'MA-734119',
    bidder: 'BD-291045',
    yard: '418'
  },
  {
    broker: 'Tri-State Dealer Services LLC',
    master: 'MA-840662',
    bidder: 'BD-175603',
    yard: '108'
  },
  {
    broker: 'Meridian Global Auto Traders',
    master: 'MA-951208',
    bidder: 'BD-663410',
    yard: '145'
  }
]);

