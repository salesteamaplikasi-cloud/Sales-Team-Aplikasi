import React, { useEffect, useRef, useState } from 'react';
import * as d3 from 'd3';

interface BulletData {
  salesmanName: string;
  actual: number;
  target: number;
  prevMonth: number;
}

export const SalesPerformanceBulletChart: React.FC<{ data: BulletData[] }> = ({ data }) => {
  const [selectedSalesman, setSelectedSalesman] = useState<string>("All");
  const chartRef = useRef<HTMLDivElement>(null);

  const filteredData = selectedSalesman === "All" 
    ? data 
    : data.filter(d => d.salesmanName === selectedSalesman);

  useEffect(() => {
    if (!chartRef.current || filteredData.length === 0) return;

    d3.select(chartRef.current).selectAll('*').remove();

    const margin = { top: 20, right: 40, bottom: 20, left: 100 };
    const width = 500 - margin.left - margin.right;
    const height = 40 - margin.top - margin.bottom;

    const svg = d3.select(chartRef.current)
      .selectAll('svg')
      .data(filteredData)
      .enter()
      .append('svg')
      .attr('width', width + margin.left + margin.right)
      .attr('height', height + margin.top + margin.bottom)
      .append('g')
      .attr('transform', `translate(${margin.left},${margin.top})`);

    const maxVal = d3.max(filteredData, d => Math.max(d.actual, d.target, d.prevMonth)) || 100;
    const x = d3.scaleLinear().domain([0, maxVal * 1.2]).range([0, width]);

    // Name
    svg.append('text')
      .attr('x', -margin.left)
      .attr('y', height / 2)
      .attr('dy', '0.35em')
      .text(d => d.salesmanName)
      .style('font-size', '12px');

    // Range (Background - placeholder for range, simplified)
    svg.append('rect')
      .attr('x', 0)
      .attr('y', 0)
      .attr('width', width)
      .attr('height', height)
      .attr('fill', '#e2e8f0');

    // Actual
    svg.append('rect')
      .attr('x', 0)
      .attr('y', height / 4)
      .attr('width', d => x(d.actual))
      .attr('height', height / 2)
      .attr('fill', d => d.actual >= d.target ? '#10b981' : '#f59e0b');

    // Target
    svg.append('line')
      .attr('x1', d => x(d.target))
      .attr('x2', d => x(d.target))
      .attr('y1', 0)
      .attr('y2', height)
      .attr('stroke', '#1e293b')
      .attr('stroke-width', 3);

    // Prev Month (marker)
    svg.append('line')
      .attr('x1', d => x(d.prevMonth))
      .attr('x2', d => x(d.prevMonth))
      .attr('y1', height / 4)
      .attr('y2', height - (height / 4))
      .attr('stroke', '#94a3b8')
      .attr('stroke-width', 2)
      .attr('stroke-dasharray', '2,2');

  }, [filteredData]);

  return (
    <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-sm">
      <div className="flex justify-between items-center mb-4">
        <h3 className="text-sm font-bold text-gray-700 uppercase">Sales Performance: Actual vs Target</h3>
        <select 
          value={selectedSalesman} 
          onChange={(e) => setSelectedSalesman(e.target.value)}
          className="text-xs border border-gray-300 rounded p-1"
        >
          <option value="All">All Salesmen</option>
          {data.map(d => (
            <option key={d.salesmanName} value={d.salesmanName}>{d.salesmanName}</option>
          ))}
        </select>
      </div>
      <div ref={chartRef} />
    </div>
  );
};
