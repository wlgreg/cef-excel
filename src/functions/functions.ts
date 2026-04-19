/* global clearInterval, console, CustomFunctions, fetch, setInterval */

/**
 * Adds two numbers.
 * @customfunction
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
export function add(first: number, second: number): number {
  return first + second;
}

/**
 * Displays the current time once a second.
 * @customfunction
 * @param invocation Custom function handler
 */
export function clock(invocation: CustomFunctions.StreamingInvocation<string>): void {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time.
 * @returns String with the current time formatted for the current locale.
 */
export function currentTime(): string {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param incrementBy Amount to increment
 * @param invocation Custom function handler
 */
export function increment(
  incrementBy: number,
  invocation: CustomFunctions.StreamingInvocation<number>
): void {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param message String to write.
 * @returns String to write.
 */
export function logMessage(message: string): string {
  console.log(message);

  return message;
}

/**
 * Gets CEF data for a ticker and endpoint.
 * @customfunction GETCEFDATA
 * @param ticker CEF ticker symbol, e.g. AWP.
 * @param endpoint Data endpoint to retrieve. Supports NAV, PRICE, DISCOUNT, DISCOUNT5YAVG.
 * @returns Requested value for the specified ticker and endpoint.
 */
export async function getCEFData(ticker: string, endpoint: string): Promise<number> {
  const normalizedTicker = (ticker || "").trim().toUpperCase();
  const normalizedEndpoint = (endpoint || "").trim().toUpperCase();

  if (!normalizedTicker) {
    throw new Error("Ticker is required.");
  }

  if (!normalizedEndpoint) {
    throw new Error("Endpoint is required.");
  }

  switch (normalizedEndpoint) {
    case "NAV":
      return await getDailyPricingValue(normalizedTicker, "NAV");
    case "PRICE":
      return await getDailyPricingValue(normalizedTicker, "PRICE");
    case "DISCOUNT":
      return await getDailyPricingValue(normalizedTicker, "DISCOUNT");
    case "DISCOUNT5YAVG":
    case "5YDISCOUNT":
      return await getFiveYearAverageDiscount(normalizedTicker);
    default:
      throw new Error(
        `Unsupported endpoint '${endpoint}'. Currently supported: NAV, PRICE, DISCOUNT, DISCOUNT5YAVG.`
      );
  }
}

interface DailyPricingRecord {
  Ticker: string;
  Price?: number;
  NAV?: number;
  Discount?: number;
}

interface DailyPricingCacheEntry {
  expiresAt: number;
  data: DailyPricingRecord[];
}

interface PricingHistoryPoint {
  DiscountData?: number;
}

interface PricingHistoryResponse {
  Data?: {
    PriceHistory?: PricingHistoryPoint[];
  };
}

const DAILY_PRICING_CACHE_TTL_MS = 10000;
let dailyPricingCache: DailyPricingCacheEntry | null = null;

async function getDailyPricingValue(
  ticker: string,
  valueType: "NAV" | "PRICE" | "DISCOUNT"
): Promise<number> {
  const data = await getDailyPricingData();
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error("DailyPricing API returned no data.");
  }

  const tickerRow = data.find((record) => String(record.Ticker || "").toUpperCase() === ticker);
  if (!tickerRow) {
    throw new Error(`Ticker '${ticker}' was not found in DailyPricing data.`);
  }

  const value =
    valueType === "PRICE"
      ? tickerRow.Price
      : valueType === "NAV"
        ? tickerRow.NAV
        : tickerRow.Discount;
  if (typeof value !== "number" || Number.isNaN(value)) {
    throw new Error(`Could not find ${valueType} value for ticker '${ticker}'.`);
  }

  return value;
}

async function getDailyPricingData(): Promise<DailyPricingRecord[]> {
  const now = Date.now();
  if (dailyPricingCache && dailyPricingCache.expiresAt > now) {
    return dailyPricingCache.data;
  }

  const url = "https://www.cefconnect.com/api/v3/DailyPricing?props=Ticker,Price,NAV,Discount";
  const response = await fetch(url, {
    method: "GET",
    headers: {
      Accept: "application/json",
    },
  });

  if (!response.ok) {
    throw new Error(`CEFConnect request failed (${response.status}).`);
  }

  const data = (await response.json()) as DailyPricingRecord[];
  if (!Array.isArray(data)) {
    throw new Error("DailyPricing API returned an unexpected response format.");
  }

  dailyPricingCache = {
    data,
    expiresAt: now + DAILY_PRICING_CACHE_TTL_MS,
  };

  return data;
}

async function getFiveYearAverageDiscount(ticker: string): Promise<number> {
  const url = `https://www.cefconnect.com/api/v3/pricinghistory/${encodeURIComponent(ticker)}/5Y`;
  const response = await fetch(url, {
    method: "GET",
    headers: {
      Accept: "application/json",
    },
  });

  if (!response.ok) {
    throw new Error(`CEFConnect request failed (${response.status}).`);
  }

  const json = (await response.json()) as PricingHistoryResponse;
  const points = json?.Data?.PriceHistory;
  const discounts = Array.isArray(points)
    ? points
        .map((point) => point?.DiscountData)
        .filter((value): value is number => typeof value === "number" && !Number.isNaN(value))
    : [];

  if (discounts.length > 0) {
    return discounts.reduce((sum, value) => sum + value, 0) / discounts.length;
  }

  return await getFiveYearAverageDiscountFromSummaryPage(ticker);
}

async function getFiveYearAverageDiscountFromSummaryPage(ticker: string): Promise<number> {
  const url = `https://www.cefconnect.com/fund/${encodeURIComponent(ticker)}`;
  const response = await fetch(url, {
    method: "GET",
    headers: {
      Accept: "text/html",
    },
  });

  if (!response.ok) {
    throw new Error(`CEFConnect request failed (${response.status}).`);
  }

  const html = await response.text();
  const discountTable = html.match(/<table[^>]*id="[^"]*DiscountGrid"[\s\S]*?<\/table>/i)?.[0];
  if (!discountTable) {
    throw new Error(`Could not find discount data for ticker '${ticker}'.`);
  }

  const fiveYearRow = discountTable.match(
    /<tr[^>]*>[\s\S]*?<td[^>]*>\s*5 Year\s*<\/td>\s*<td[^>]*>([\s\S]*?)<\/td>[\s\S]*?<\/tr>/i
  )?.[1];
  if (!fiveYearRow) {
    throw new Error(`Could not find 5 Year average discount for ticker '${ticker}'.`);
  }

  const parsedValue = parseNumberFromHtmlCell(fiveYearRow);
  if (Number.isNaN(parsedValue)) {
    throw new Error(`Unable to parse 5 Year average discount for ticker '${ticker}'.`);
  }

  return parsedValue;
}

function parseNumberFromHtmlCell(cellContent: string): number {
  const withoutTags = cellContent.replace(/<[^>]*>/g, "");
  const normalized = withoutTags
    .replace(/&nbsp;/gi, " ")
    .replace(/\$/g, "")
    .replace(/%/g, "")
    .replace(/,/g, "")
    .trim();

  return Number.parseFloat(normalized);
}
