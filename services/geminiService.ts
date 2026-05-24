import { GoogleGenAI } from "@google/genai";
import { DealerData } from '../types';

export const getGeminiInsights = async (data: DealerData[], promptContext: string): Promise<string> => {
  const apiKey = process.env.API_KEY;
  if (!apiKey) {
    throw new Error("API Key not found in environment variables.");
  }

  const ai = new GoogleGenAI({ apiKey });

  // Send a larger slice (200 records) to ensure the analysis is based on "real" data 
  // rather than a tiny sample. Gemini Flash has a large context window.
  const jsonContext = JSON.stringify(data.slice(0, 200)); 

  const prompt = `
    You are a precision data analyst. 
    Your task is to answer the User Query based ONLY on the provided JSON dataset.
    
    Rules:
    1. Be extremely concise, concrete, and factual.
    2. Do not use filler words like "Based on the data provided" or "Here is the analysis".
    3. Cite specific numbers, dealer names, models, or percentages from the data to support your answer.
    4. If the user asks for general insights, provide exactly 3 short, data-backed bullet points.
    5. If the user asks a specific question, answer it directly with data.
    
    User Query: ${promptContext}
    
    Dataset (First 200 rows):
    ${jsonContext}
  `;

  try {
    const response = await ai.models.generateContent({
      model: 'gemini-3-flash-preview',
      contents: prompt,
      config: {
        thinkingConfig: { thinkingBudget: 0 } // Disable thinking for faster response
      }
    });
    
    return response.text || "No insights generated.";
  } catch (error) {
    console.error("Gemini API Error:", error);
    return "Failed to generate insights. Please check your API key or try again later.";
  }
};