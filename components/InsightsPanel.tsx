import React, { useState } from 'react';
import { DealerData } from '../types';
import { getGeminiInsights } from '../services/geminiService';
import { Sparkles, Send, Loader2, Bot } from 'lucide-react';

interface InsightsPanelProps {
  data: DealerData[];
}

export const InsightsPanel: React.FC<InsightsPanelProps> = ({ data }) => {
  const [query, setQuery] = useState('');
  const [response, setResponse] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);

  const handleGenerate = async () => {
    if (!query.trim()) return;
    setLoading(true);
    setResponse(null);
    
    try {
      const result = await getGeminiInsights(data, query);
      setResponse(result);
    } catch (e) {
      setResponse("An error occurred while communicating with Gemini.");
    } finally {
      setLoading(false);
    }
  };

  const suggestions = [
    "Identify the top performing dealers by Sales Volume",
    "Which region has the lowest Customer Satisfaction?",
    "Summarize the compliance status across all dealers",
    "Find correlation between Showroom Size and Sales"
  ];

  return (
    <div className="bg-white rounded-xl shadow-sm border border-indigo-100 overflow-hidden flex flex-col h-full">
      <div className="p-6 bg-gradient-to-r from-indigo-600 to-blue-600 text-white">
        <h2 className="text-xl font-bold flex items-center gap-2">
          <Sparkles className="text-yellow-300" size={24} />
          AI Data Analyst
        </h2>
        <p className="text-indigo-100 mt-2 text-sm opacity-90">
          Ask questions about your dealer network data and get instant, intelligent analysis powered by Gemini.
        </p>
      </div>

      <div className="p-6 flex-1 overflow-y-auto bg-slate-50/50">
        {!response && !loading && (
          <div className="space-y-4">
            <p className="text-sm font-medium text-slate-500 uppercase tracking-wide">Suggested Queries</p>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
              {suggestions.map((s, i) => (
                <button 
                  key={i}
                  onClick={() => setQuery(s)}
                  className="text-left p-3 rounded-lg border border-slate-200 bg-white hover:border-indigo-300 hover:shadow-sm transition-all text-sm text-slate-700"
                >
                  {s}
                </button>
              ))}
            </div>
          </div>
        )}

        {loading && (
          <div className="flex flex-col items-center justify-center h-48 space-y-4">
            <Loader2 className="animate-spin text-indigo-600" size={40} />
            <p className="text-slate-500 text-sm animate-pulse">Analyzing dealer data patterns...</p>
          </div>
        )}

        {response && (
          <div className="prose prose-sm max-w-none prose-indigo bg-white p-6 rounded-xl border border-slate-200 shadow-sm">
            <div className="flex items-center gap-2 mb-4 text-indigo-700 font-semibold border-b border-slate-100 pb-2">
              <Bot size={20} />
              Analysis Result
            </div>
             {/* Simple markdown rendering for the demo */}
            <div className="whitespace-pre-wrap text-slate-700 leading-relaxed">
              {response}
            </div>
          </div>
        )}
      </div>

      <div className="p-4 bg-white border-t border-slate-200">
        <div className="relative">
          <input
            type="text"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            onKeyDown={(e) => e.key === 'Enter' && handleGenerate()}
            placeholder="Ask a question about the data..."
            className="w-full pl-4 pr-12 py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none shadow-sm"
          />
          <button
            onClick={handleGenerate}
            disabled={loading || !query.trim()}
            className="absolute right-2 top-1/2 -translate-y-1/2 p-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
          >
            <Send size={16} />
          </button>
        </div>
      </div>
    </div>
  );
};