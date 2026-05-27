import React, { useEffect, useRef } from 'react';
import * as d3 from 'd3';
import { NooRecord } from '../types';

export const NooSummary: React.FC<{ records: NooRecord[] }> = ({ records }) => {
  const svgRef = useRef<SVGSVGElement>(null);

  useEffect(() => {
    if (!svgRef.current || records.length === 0) return;

    // Aggregate data
    const summary = records.reduce((acc, curr) => {
      acc.warung = (acc.warung || 0) + (curr.warung || 0);
      acc.store = (acc.store || 0) + (curr.store || 0);
      acc.kiosk = (acc.kiosk || 0) + (curr.kiosk || 0);
      acc.wholesaler = (acc.wholesaler || 0) + (curr.wholesaler || 0);
      return acc;
    }, { warung: 0, store: 0, kiosk: 0, wholesaler: 0 } as any);

    const data = Object.entries(summary).map(([name, value]) => ({ name, value: value as number }));

    // Set dimensions
    const width = 400;
    const height = 200;
    const margin = { top: 20, right: 20, bottom: 30, left: 40 };

    const svg = d3.select(svgRef.current);
    svg.selectAll('*').remove();

    const x = d3.scaleBand()
      .domain(data.map(d => d.name))
      .range([margin.left, width - margin.right])
      .padding(0.1);

    const y = d3.scaleLinear()
      .domain([0, d3.max(data, d => d.value) || 0])
      .range([height - margin.bottom, margin.top]);

    svg.append('g')
      .selectAll('rect')
      .data(data)
      .enter()
      .append('rect')
      .attr('x', d => x(d.name) || 0)
      .attr('y', d => y(d.value))
      .attr('width', x.bandwidth())
      .attr('height', d => height - margin.bottom - y(d.value))
      .attr('fill', '#f59e0b');

    svg.append('g')
      .attr('transform', `translate(0, ${height - margin.bottom})`)
      .call(d3.axisBottom(x));

    svg.append('g')
      .attr('transform', `translate(${margin.left}, 0)`)
      .call(d3.axisLeft(y).ticks(5));

  }, [records]);

  return (
    <div className="bg-white p-4 rounded-xl border border-gray-200 shadow-sm">
      <h3 className="text-sm font-bold text-gray-700 mb-4 uppercase">Tren NOO Berdasarkan Kategori</h3>
      <svg ref={svgRef} width="400" height="200" />
    </div>
  );
};
