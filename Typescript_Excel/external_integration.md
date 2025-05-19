# TypeScript Integration with Excel: External APIs and Tools

Excel has evolved from a simple spreadsheet application to a powerful platform that can integrate with external systems through TypeScript and JavaScript. Let me walk you through the key concepts and provide some practical examples.

## Key Concepts

1. **Office Add-ins Framework**: Microsoft's platform for building solutions that extend Office applications
2. **Excel JavaScript API**: Provides TypeScript-compatible interfaces to interact with Excel objects
3. **External API Integration**: Connecting Excel to web services and data sources
4. **Development Tools**: Modern tooling for TypeScript-based Excel development

## Office Add-ins and TypeScript

Office Add-ins use web technologies (HTML, CSS, JavaScript/TypeScript) that run in a secure container within Excel. TypeScript provides strong typing, which is especially valuable when working with the Excel object model.

### Example: Basic Excel Add-in Setup

```typescript
// Excel Add-in TypeScript entry point
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Initialize your Excel add-in
    console.log("Excel add-in initialized");
    
    // Register event handlers or UI elements
    document.getElementById("run-button").onclick = runOperation;
  }
});

async function runOperation() {
  try {
    await Excel.run(async (context) => {
      // Get the current worksheet
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      // Define a range
      const range = sheet.getRange("A1:B5");
      range.format.fill.color = "yellow";
      
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
}
```

## Integrating External APIs

TypeScript makes it easier to work with external APIs by providing type safety for request and response data.

### Example: Fetching Stock Data

```typescript
interface StockData {
  symbol: string;
  price: number;
  change: number;
  changePercent: number;
}

async function getStockData(symbol: string): Promise<StockData> {
  const response = await fetch(`https://api.example.com/stocks/${symbol}`);
  
  if (!response.ok) {
    throw new Error(`API error: ${response.status}`);
  }
  
  return await response.json() as StockData;
}

async function updateStockInfo() {
  await Excel.run(async (context) => {
    try {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      // Get stock symbols from column A
      const symbolRange = sheet.getRange("A2:A10");
      symbolRange.load("values");
      await context.sync();
      
      const symbols = symbolRange.values.flat().filter(Boolean);
      
      // Fetch data for each symbol
      for (let i = 0; i < symbols.length; i++) {
        const symbol = symbols[i];
        const data = await getStockData(symbol);
        
        // Update cells with stock data
        sheet.getRange(`B${i+2}`).values = [[data.price]];
        sheet.getRange(`C${i+2}`).values = [[data.change]];
        sheet.getRange(`D${i+2}`).values = [[data.changePercent]];
        
        // Format cells based on values
        if (data.change < 0) {
          sheet.getRange(`C${i+2}:D${i+2}`).format.font.color = "red";
        } else {
          sheet.getRange(`C${i+2}:D${i+2}`).format.font.color = "green";
        }
      }
      
      await context.sync();
    } catch (error) {
      console.error("Error updating stock data:", error);
    }
  });
}
```

## Custom Functions with TypeScript

Excel custom functions let you define your own spreadsheet functions that can call external APIs.

### Example: Currency Conversion Function

```typescript
interface ExchangeRateResponse {
  rates: {
    [currency: string]: number;
  };
  base: string;
}

/**
 * Converts an amount from one currency to another
 * @customfunction
 * @param {number} amount The amount to convert
 * @param {string} fromCurrency The source currency code
 * @param {string} toCurrency The target currency code
 * @returns {number} The converted amount
 */
async function CURRENCY_CONVERT(amount: number, fromCurrency: string, toCurrency: string): Promise<number> {
  try {
    const response = await fetch(`https://api.exchangerate-api.com/v4/latest/${fromCurrency}`);
    const data = await response.json() as ExchangeRateResponse;
    
    if (!data.rates[toCurrency]) {
      throw new Error(`Invalid currency code: ${toCurrency}`);
    }
    
    return amount * data.rates[toCurrency];
  } catch (error) {
    throw new Error(`Conversion error: ${error.message}`);
  }
}

// Register the custom function
CustomFunctions.associate("CURRENCY_CONVERT", CURRENCY_CONVERT);
```

## Integration with Modern Development Tools

TypeScript-based Excel add-ins work well with modern development tools and frameworks.

### Example: React + TypeScript Excel Add-in

```typescript
import * as React from 'react';
import { useState, useEffect } from 'react';
import { DefaultButton } from '@fluentui/react';

interface WeatherData {
  temperature: number;
  description: string;
  humidity: number;
}

const WeatherComponent: React.FC = () => {
  const [weather, setWeather] = useState<WeatherData | null>(null);
  const [loading, setLoading] = useState<boolean>(false);
  const [location, setLocation] = useState<string>('');
  
  const insertWeatherData = async () => {
    try {
      setLoading(true);
      
      // Get weather data from API
      const response = await fetch(`https://api.weatherapi.com/v1/current.json?key=YOUR_API_KEY&q=${location}`);
      const data = await response.json();
      
      const weatherData: WeatherData = {
        temperature: data.current.temp_c,
        description: data.current.condition.text,
        humidity: data.current.humidity
      };
      
      setWeather(weatherData);
      
      // Insert weather data into Excel
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        
        sheet.getRange("A1").values = [["Location"]];
        sheet.getRange("A2").values = [[location]];
        
        sheet.getRange("B1").values = [["Temperature (°C)"]];
        sheet.getRange("B2").values = [[weatherData.temperature]];
        
        sheet.getRange("C1").values = [["Description"]];
        sheet.getRange("C2").values = [[weatherData.description]];
        
        sheet.getRange("D1").values = [["Humidity (%)"]];
        sheet.getRange("D2").values = [[weatherData.humidity]];
        
        sheet.getRange("A1:D1").format.font.bold = true;
        
        await context.sync();
      });
    } catch (error) {
      console.error("Error fetching weather data:", error);
    } finally {
      setLoading(false);
    }
  };
  
  return (
    <div>
      <input 
        type="text" 
        value={location}
        onChange={(e) => setLocation(e.target.value)}
        placeholder="Enter location"
      />
      <DefaultButton 
        text={loading ? "Loading..." : "Get Weather"}
        disabled={loading || !location}
        onClick={insertWeatherData}
      />
      {weather && (
        <div>
          <p>Temperature: {weather.temperature}°C</p>
          <p>Description: {weather.description}</p>
          <p>Humidity: {weather.humidity}%</p>
        </div>
      )}
    </div>
  );
};

export default WeatherComponent;
```

## Power Query Integration

TypeScript can be used to enhance the Power Query M experience with custom connectors:

```typescript
// TypeScript definition for Power Query connector
interface PowerQueryConnector {
  name: string;
  dataSourceKind: string;
  dataSources: {
    name: string;
    connectionString: string;
  }[];
}

// Generate Power Query M code from TypeScript
function generatePowerQueryConnector(config: PowerQueryConnector): string {
  return `
    [Version = "1.0.0"]
    section ${config.name};
    
    [DataSource.Kind="${config.dataSourceKind}"]
    shared ${config.name}.Contents = (url as text) =>
      let
        source = Web.Contents(url),
        json = Json.Document(source)
      in
        json;
    
    // Data Source UI publishing description
    ${config.name}.Publish = [
      Beta = true,
      ButtonText = { "Connect to ${config.name}", "Connecting to ${config.name}" },
      SourceImage = ${config.name}.Icons,
      SourceTypeImage = ${config.name}.Icons
    ];
    
    ${config.name}.Icons = [
      Icon16 = { Extension.Contents("${config.name}16.png"), Extension.Contents("${config.name}20.png"), Extension.Contents("${config.name}24.png"), Extension.Contents("${config.name}32.png") },
      Icon32 = { Extension.Contents("${config.name}32.png"), Extension.Contents("${config.name}40.png"), Extension.Contents("${config.name}48.png"), Extension.Contents("${config.name}64.png") }
    ];
  `;
}
```

## Additional Integration Examples

### Example: Real-time Data Dashboard with SignalR

```typescript
import * as signalR from '@microsoft/signalr';

interface DataPoint {
  timestamp: Date;
  value: number;
}

class RealTimeDataService {
  private connection: signalR.HubConnection;
  private dataPoints: DataPoint[] = [];
  
  constructor(hubUrl: string) {
    this.connection = new signalR.HubConnectionBuilder()
      .withUrl(hubUrl)
      .withAutomaticReconnect()
      .build();
      
    this.connection.on("newDataPoint", (data: DataPoint) => {
      this.dataPoints.push(data);
      this.updateExcel();
    });
  }
  
  async start(): Promise<void> {
    try {
      await this.connection.start();
      console.log("Connected to SignalR hub");
    } catch (error) {
      console.error("Error connecting to SignalR hub:", error);
      setTimeout(() => this.start(), 5000);
    }
  }
  
  private async updateExcel(): Promise<void> {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      // Clear existing data
      sheet.getRange("A2:B100").clear();
      
      // Add latest data points (up to 50)
      const dataToDisplay = this.dataPoints.slice(-50);
      
      for (let i = 0; i < dataToDisplay.length; i++) {
        const row = i + 2; // Starting at row 2
        sheet.getRange(`A${row}`).values = [[dataToDisplay[i].timestamp.toLocaleString()]];
        sheet.getRange(`B${row}`).values = [[dataToDisplay[i].value]];
      }
      
      // Create or update chart
      const charts = sheet.charts;
      let chart = charts.getItemOrNullObject("DataChart");
      await context.sync();
      
      if (chart.isNullObject) {
        chart = charts.add(Excel.ChartType.line, sheet.getRange("A2:B" + (dataToDisplay.length + 1)), Excel.ChartSeriesBy.columns);
        chart.name = "DataChart";
        chart.title.text = "Real-time Data";
      } else {
        chart.setData(sheet.getRange("A2:B" + (dataToDisplay.length + 1)), Excel.ChartSeriesBy.columns);
      }
      
      await context.sync();
    });
  }
}

// Usage
Office.onReady(async () => {
  const dataService = new RealTimeDataService("https://your-signalr-hub.com/datahub");
  await dataService.start();
});
```

These examples demonstrate how TypeScript can be used to create robust, type-safe integrations between Excel and external APIs and tools, enhancing the functionality of Excel while leveraging modern development practices.
