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
 * @param endpoint Data endpoint to retrieve. Supports NAV, PRICE, DISCOUNT, DISCOUNT5YAVG, DISTYIELDNAV, DISTYIELDPRICE, ANNDISTRATENAV1Y, ANNDISTRATENAV3Y, ANNDISTRATENAV5Y.
 * @param debug Optional TRUE to return diagnostic text instead of #VALUE on errors.
 * @returns Requested value for the specified ticker and endpoint, or N/A when no ticker data is available.
 */
export async function getCEFData(
  ticker: string,
  endpoint: string,
  debug?: boolean
): Promise<number | string> {
  const normalizedTicker = (ticker || "").trim().toUpperCase();
  const normalizedEndpoint = (endpoint || "").trim().toUpperCase();

  try {
    if (!normalizedTicker) {
      throw new Error("Ticker is required.");
    }

    if (!normalizedEndpoint) {
      throw new Error("Endpoint is required.");
    }

    switch (normalizedEndpoint) {
      case "NAV":
        return (await getDailyPricingValue(normalizedTicker, "NAV")) ?? "N/A";
      case "PRICE":
        return (await getDailyPricingValue(normalizedTicker, "PRICE")) ?? "N/A";
      case "DISCOUNT":
        return (await getDailyPricingValue(normalizedTicker, "DISCOUNT")) ?? "N/A";
      case "DISTYIELDNAV":
      case "YIELDNAV":
        return (await getDailyPricingValue(normalizedTicker, "DISTYIELDNAV")) ?? "N/A";
      case "DISTYIELDPRICE":
      case "YIELDPRICE":
        return (await getDailyPricingValue(normalizedTicker, "DISTYIELDPRICE")) ?? "N/A";
      case "ANNDISTRATENAV1Y":
      case "DISTYIELDNAV1Y":
        return (await getAnnualizedDistributionRateOnNav(normalizedTicker, 1)) ?? "N/A";
      case "ANNDISTRATENAV3Y":
      case "DISTYIELDNAV3Y":
        return (await getAnnualizedDistributionRateOnNav(normalizedTicker, 3)) ?? "N/A";
      case "ANNDISTRATENAV5Y":
      case "DISTYIELDNAV5Y":
        return (await getAnnualizedDistributionRateOnNav(normalizedTicker, 5)) ?? "N/A";
      case "DISCOUNT5YAVG":
      case "5YDISCOUNT":
        return (await getFiveYearAverageDiscount(normalizedTicker)) ?? "N/A";
      default:
        throw new Error(
          `Unsupported endpoint '${endpoint}'. Currently supported: NAV, PRICE, DISCOUNT, DISCOUNT5YAVG, DISTYIELDNAV (alias YIELDNAV), DISTYIELDPRICE (alias YIELDPRICE), ANNDISTRATENAV1Y, ANNDISTRATENAV3Y, ANNDISTRATENAV5Y.`
        );
    }
  } catch (error) {
    if (debug === true) {
      const message = error instanceof Error ? error.message : String(error);
      return `ERROR | ticker=${normalizedTicker || "(blank)"} | endpoint=${normalizedEndpoint || "(blank)"} | message=${message} | cache=${getDailyPricingCacheStatus()} | time=${new Date().toISOString()}`;
    }

    throw error;
  }
}

/**
 * Returns diagnostic information for a GETCEFDATA request.
 * @customfunction GETCEFDATADEBUG
 * @param ticker CEF ticker symbol, e.g. AWP.
 * @param endpoint Data endpoint to retrieve.
 * @returns Diagnostic status text and either value or detailed error information.
 */
export async function getCEFDataDebug(ticker: string, endpoint: string): Promise<string> {
  const normalizedTicker = (ticker || "").trim().toUpperCase();
  const normalizedEndpoint = (endpoint || "").trim().toUpperCase();

  try {
    const value = await getCEFData(normalizedTicker, normalizedEndpoint);
    return `OK | ticker=${normalizedTicker || "(blank)"} | endpoint=${normalizedEndpoint || "(blank)"} | value=${value} | cache=${getDailyPricingCacheStatus()} | time=${new Date().toISOString()}`;
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return `ERROR | ticker=${normalizedTicker || "(blank)"} | endpoint=${normalizedEndpoint || "(blank)"} | message=${message} | cache=${getDailyPricingCacheStatus()} | time=${new Date().toISOString()}`;
  }
}

interface DailyPricingRecord {
  Ticker: string;
  Price?: number;
  NAV?: number;
  Discount?: number;
  DistributionRateNAV?: number;
  DistributionRatePrice?: number;
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

interface DistributionCharterPoint {
  Date?: string;
  Amount?: number;
}

interface DistributionCharterResponse {
  Data?: DistributionCharterPoint[];
}

const DAILY_PRICING_CACHE_TTL_MS = 10000;
let dailyPricingCache: DailyPricingCacheEntry | null = null;

function getDailyPricingCacheStatus(): string {
  if (!dailyPricingCache) {
    return "empty";
  }

  const msToExpiry = dailyPricingCache.expiresAt - Date.now();
  if (msToExpiry > 0) {
    return `warm:${Math.floor(msToExpiry / 1000)}s`;
  }

  return `stale:${Math.abs(Math.floor(msToExpiry / 1000))}s`;
}

async function getDailyPricingValue(
  ticker: string,
  valueType: "NAV" | "PRICE" | "DISCOUNT" | "DISTYIELDNAV" | "DISTYIELDPRICE"
): Promise<number | null> {
  const data = await getDailyPricingData();
  if (!Array.isArray(data) || data.length === 0) {
    throw new Error("DailyPricing API returned no data.");
  }

  const tickerRow = data.find((record) => String(record.Ticker || "").toUpperCase() === ticker);
  if (!tickerRow) {
    return null;
  }

  const value =
    valueType === "PRICE"
      ? tickerRow.Price
      : valueType === "NAV"
        ? tickerRow.NAV
        : valueType === "DISCOUNT"
          ? tickerRow.Discount
          : valueType === "DISTYIELDNAV"
            ? tickerRow.DistributionRateNAV
            : tickerRow.DistributionRatePrice;
  if (typeof value !== "number" || Number.isNaN(value)) {
    return null;
  }

  return value;
}

async function getDailyPricingData(): Promise<DailyPricingRecord[]> {
  const now = Date.now();
  if (dailyPricingCache && dailyPricingCache.expiresAt > now) {
    return dailyPricingCache.data;
  }

  const url =
    "https://www.cefconnect.com/api/v3/DailyPricing?props=Ticker,Price,NAV,Discount,DistributionRateNAV,DistributionRatePrice";
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

async function getAnnualizedDistributionRateOnNav(
  ticker: string,
  years: 1 | 3 | 5
): Promise<number | null> {
  const nav = await getDailyPricingValue(ticker, "NAV");
  if (typeof nav !== "number" || Number.isNaN(nav) || nav <= 0) {
    return null;
  }

  const period = `${years}Y`;
  const url = `https://www.cefconnect.com/api/v3/DistributionCharter/fund/${encodeURIComponent(ticker)}/${period}`;
  const response = await fetch(url, {
    method: "GET",
    headers: {
      Accept: "application/json",
    },
  });

  if (!response.ok) {
    throw new Error(`CEFConnect request failed (${response.status}).`);
  }

  const json = (await response.json()) as DistributionCharterResponse;
  const points = Array.isArray(json?.Data) ? json.Data : [];
  if (points.length === 0) {
    return null;
  }

  const now = Date.now();
  const totalDistributions = points.reduce((sum, point) => {
    const amount = point?.Amount;
    const dateMs = point?.Date ? Date.parse(point.Date) : Number.NaN;
    if (typeof amount !== "number" || Number.isNaN(amount)) {
      return sum;
    }

    // Ignore future dated payouts when computing trailing annualized rates.
    if (!Number.isNaN(dateMs) && dateMs > now) {
      return sum;
    }

    return sum + amount;
  }, 0);

  if (totalDistributions <= 0) {
    return null;
  }

  const annualizedAmount = totalDistributions / years;
  return (annualizedAmount / nav) * 100;
}

async function getFiveYearAverageDiscount(ticker: string): Promise<number | null> {
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

async function getFiveYearAverageDiscountFromSummaryPage(ticker: string): Promise<number | null> {
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
    return null;
  }

  const fiveYearRow = discountTable.match(
    /<tr[^>]*>[\s\S]*?<td[^>]*>\s*5 Year\s*<\/td>\s*<td[^>]*>([\s\S]*?)<\/td>[\s\S]*?<\/tr>/i
  )?.[1];
  if (!fiveYearRow) {
    return null;
  }

  const parsedValue = parseNumberFromHtmlCell(fiveYearRow);
  if (Number.isNaN(parsedValue)) {
    return null;
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
