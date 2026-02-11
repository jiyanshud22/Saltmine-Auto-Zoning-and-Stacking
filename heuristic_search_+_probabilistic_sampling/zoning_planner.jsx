import React, { useState, useEffect, useRef } from 'react';
import { AlertCircle, CheckCircle, XCircle, Info } from 'lucide-react';

// ============ DATA GENERATION & MANAGEMENT ============

const BLOCK_TYPES = {
  ME: ['Individual Zone Work Desks (no partition)', 'Individual Zone Focus Rooms (1p)', 'Individual Zone Focus Rooms (2-3p)', 
       'Interactive Zone Work Desks (with partitions)', 'Interactive Zone Focus Rooms (1p)', 'Interactive Zone Focus Rooms (2-3p)',
       'Individual Zone Touchdown Seats', 'Interactive Zone Touchdown Seats'],
  WE: ['Individual Zone Phone Booths (Single)', 'Individual Zone Phone Booths (Double)', 'Interactive Zone Phone Booths (Single)',
       'Interactive Zone Phone Booths (Double)', 'Huddle Room (4-6 pax)', 'Small Meeting (7-9 pax)', 'Medium Meeting (10-12 pax)',
       'Large Meeting/Boardroom (13-30 pax)'],
  US: ['Restaurant (full-service café)', 'Coffee Bar (barista-serviced coffee shop)', 'Alternative Food Points', 
       'Flex Space (formerly Agile Space)', 'Visitor Touchdown'],
  SUPPORT: ['Lockers', 'Team Storage', 'Copy/Print', 'General Storage', 'Mail/Shipping Center', 
            'Kitchen/Pantry (with self-service beverage)', 'Waste Management', 'Janitor\'s Closet'],
  SPECIALTY: ['IDF', 'Server Room', 'Security Room', 'Accessibility Lab']
};

const BUILDING_FEATURES = {
  CORES: ['Core-A', 'Core-B', 'Core-C'],
  ANCHORS: ['Main Entry', 'Service Entry', 'Fire Egress'],
  SPECIAL: ['AHU', 'Window/Glazing Wall', 'Toilets']
};

// Priority-based adjacency rules
const ADJACENCY_RULES = {
  // Priority 0: Non-compromisable adjacencies
  0: [
    { from: 'IDF', to: 'AHU', relation: 'must_avoid', weight: -2 },
    { from: 'Server Room', to: 'AHU', relation: 'must_avoid', weight: -2 },
    { from: 'Security Room', to: 'AHU', relation: 'must_avoid', weight: -2 },
    { from: 'IDF', to: 'Window/Glazing Wall', relation: 'must_avoid', weight: -2 },
    { from: 'Server Room', to: 'Window/Glazing Wall', relation: 'must_avoid', weight: -2 },
    { from: 'IDF', to: 'Toilets', relation: 'must_avoid', weight: -2 },
    { from: 'Server Room', to: 'Toilets', relation: 'must_avoid', weight: -2 },
    { from: 'Restaurant (full-service café)', to: 'Main Entry', relation: 'must_be_near', weight: 1 },
    { from: 'Bike Storage/Lockers', to: 'Service Entry', relation: 'must_be_near', weight: 1 },
  ],
  // Priority 1: Geometrical Resolution (rectangular shapes preferred)
  1: [],
  // Priority 2: Block to Building Feature
  2: [
    { from: 'IDF', to: 'Core', relation: 'preferred_near', weight: 0.75 },
    { from: 'Server Room', to: 'Core', relation: 'preferred_near', weight: 0.75 },
    { from: 'Restaurant (full-service café)', to: 'Window/Glazing Wall', relation: 'preferred_near', weight: 0.75 },
    { from: 'Coffee Bar (barista-serviced coffee shop)', to: 'Window/Glazing Wall', relation: 'preferred_near', weight: 0.75 },
  ],
  // Priority 3: Block to Block
  3: [
    { from: 'Individual Zone Phone Booths (Single)', to: 'Individual Zone Work Desks (no partition)', relation: 'should_avoid', weight: -1 },
    { from: 'Lockers', to: 'Individual Zone Work Desks (no partition)', relation: 'should_avoid', weight: -1 },
  ],
  // Priority 4: Superblock to Superblock
  4: [],
  // Priority 5: Block to Department
  5: []
};

// Synthetic data generator
const generateSyntheticBlocks = (count = 100) => {
  const blocks = [];
  let blockId = 1;
  
  const categories = Object.keys(BLOCK_TYPES);
  
  for (let i = 0; i < count; i++) {
    const category = categories[Math.floor(Math.random() * categories.length)];
    const blockTypes = BLOCK_TYPES[category];
    const blockType = blockTypes[Math.floor(Math.random() * blockTypes.length)];
    
    // Determine dimensions based on block type
    let width, height, area;
    if (category === 'ME') {
      width = Math.random() < 0.7 ? 3 : 4; // 3 or 4 meters
      height = Math.random() < 0.7 ? 3 : 4;
      area = width * height;
    } else if (category === 'WE') {
      width = Math.random() < 0.5 ? 2 : 3;
      height = Math.random() < 0.5 ? 3 : 4;
      area = width * height;
    } else if (category === 'US') {
      width = Math.floor(Math.random() * 5) + 5; // 5-10 meters
      height = Math.floor(Math.random() * 5) + 5;
      area = width * height;
    } else if (category === 'SUPPORT') {
      width = Math.random() < 0.5 ? 2 : 3;
      height = Math.random() < 0.5 ? 2 : 3;
      area = width * height;
    } else {
      width = 3;
      height = 3;
      area = 9;
    }
    
    blocks.push({
      id: `BLK-${String(blockId).padStart(3, '0')}`,
      name: blockType,
      category,
      width,
      height,
      area,
      organization: 'Anchor Resident',
      superblockId: null,
      placed: false,
      x: null,
      y: null
    });
    blockId++;
  }
  
  return blocks;
};

// Superblock generator
const generateSuperblocks = (blocks) => {
  const superblocks = [];
  const targetCompositions = [
    { me: 0.60, we: 0.25, us: 0.10, support: 0.05, specialty: 0.0 },
    { me: 0.50, we: 0.30, us: 0.15, support: 0.05, specialty: 0.0 },
    { me: 0.55, we: 0.28, us: 0.12, support: 0.05, specialty: 0.0 },
    { me: 0.65, we: 0.20, us: 0.10, support: 0.05, specialty: 0.0 },
  ];
  
  let unassignedBlocks = [...blocks];
  let superblockId = 1;
  
  while (unassignedBlocks.length > 10 && superblockId <= 12) {
    const targetSize = Math.floor(Math.random() * 11) + 20; // 20-30 blocks
    const targetComp = targetCompositions[Math.floor(Math.random() * targetCompositions.length)];
    
    const superblockBlocks = [];
    const needed = {
      ME: Math.floor(targetSize * targetComp.me),
      WE: Math.floor(targetSize * targetComp.we),
      US: Math.floor(targetSize * targetComp.us),
      SUPPORT: Math.floor(targetSize * targetComp.support),
      SPECIALTY: Math.floor(targetSize * targetComp.specialty)
    };
    
    // Select blocks for each category
    for (const [category, count] of Object.entries(needed)) {
      const categoryBlocks = unassignedBlocks.filter(b => b.category === category);
      const selected = categoryBlocks.slice(0, Math.min(count, categoryBlocks.length));
      selected.forEach(block => {
        block.superblockId = `SB-${String(superblockId).padStart(2, '0')}`;
        superblockBlocks.push(block);
      });
      unassignedBlocks = unassignedBlocks.filter(b => !selected.includes(b));
    }
    
    if (superblockBlocks.length > 0) {
      const actualComp = {
        ME: superblockBlocks.filter(b => b.category === 'ME').length / superblockBlocks.length,
        WE: superblockBlocks.filter(b => b.category === 'WE').length / superblockBlocks.length,
        US: superblockBlocks.filter(b => b.category === 'US').length / superblockBlocks.length,
        SUPPORT: superblockBlocks.filter(b => b.category === 'SUPPORT').length / superblockBlocks.length,
        SPECIALTY: superblockBlocks.filter(b => b.category === 'SPECIALTY').length / superblockBlocks.length,
      };
      
      superblocks.push({
        id: `SB-${String(superblockId).padStart(2, '0')}`,
        blocks: superblockBlocks,
        targetComposition: targetComp,
        actualComposition: actualComp,
        compositionMatch: calculateCompositionMatch(targetComp, actualComp)
      });
    }
    
    superblockId++;
  }
  
  return superblocks;
};

const calculateCompositionMatch = (target, actual) => {
  const categories = ['ME', 'WE', 'US', 'SUPPORT', 'SPECIALTY'];
  let totalError = 0;
  categories.forEach(cat => {
    const error = Math.abs((target[cat.toLowerCase()] || 0) - (actual[cat] || 0));
    totalError += error;
  });
  return Math.max(0, 1 - totalError) * 100;
};

// ============ PLACEMENT ALGORITHM ============

const placeBlocksOnGrid = (blocks, superblocks, gridWidth = 100, gridHeight = 80) => {
  const grid = Array(gridHeight).fill(null).map(() => Array(gridWidth).fill(null));
  const placedBlocks = [];
  
  // Place building features first (cores and anchors)
  const cores = [
    { id: 'CORE-A', x: 10, y: 10, width: 5, height: 5, type: 'core' },
    { id: 'CORE-B', x: 50, y: 10, width: 5, height: 5, type: 'core' },
    { id: 'CORE-C', x: 85, y: 40, width: 5, height: 5, type: 'core' }
  ];
  
  const anchors = [
    { id: 'MAIN-ENTRY', x: 45, y: 0, width: 8, height: 3, type: 'anchor' },
    { id: 'SERVICE-ENTRY', x: 0, y: 40, width: 3, height: 5, type: 'anchor' },
    { id: 'FIRE-EGRESS-1', x: 95, y: 20, width: 3, height: 4, type: 'anchor' },
    { id: 'FIRE-EGRESS-2', x: 95, y: 60, width: 3, height: 4, type: 'anchor' }
  ];
  
  const buildingFeatures = [...cores, ...anchors];
  
  // Mark building features on grid
  buildingFeatures.forEach(feature => {
    for (let y = feature.y; y < feature.y + feature.height && y < gridHeight; y++) {
      for (let x = feature.x; x < feature.x + feature.width && x < gridWidth; x++) {
        grid[y][x] = feature.id;
      }
    }
  });
  
  // Place blocks by superblock with adjacency consideration
  superblocks.forEach((superblock, sbIndex) => {
    const startX = (sbIndex % 3) * 30 + 5;
    const startY = Math.floor(sbIndex / 3) * 25 + 15;
    
    let currentX = startX;
    let currentY = startY;
    let maxRowHeight = 0;
    
    // Sort blocks by priority (ME first, then WE, etc.)
    const sortedBlocks = [...superblock.blocks].sort((a, b) => {
      const order = ['ME', 'WE', 'US', 'SUPPORT', 'SPECIALTY'];
      return order.indexOf(a.category) - order.indexOf(b.category);
    });
    
    sortedBlocks.forEach(block => {
      let placed = false;
      let attempts = 0;
      
      while (!placed && attempts < 50) {
        // Check if block fits at current position
        if (currentX + block.width > gridWidth - 5) {
          currentX = startX;
          currentY += maxRowHeight + 1;
          maxRowHeight = 0;
        }
        
        if (currentY + block.height > gridHeight - 5) {
          // Move to next superblock area
          break;
        }
        
        // Check if space is free
        let canPlace = true;
        for (let y = currentY; y < currentY + block.height && y < gridHeight; y++) {
          for (let x = currentX; x < currentX + block.width && x < gridWidth; x++) {
            if (grid[y][x] !== null) {
              canPlace = false;
              break;
            }
          }
          if (!canPlace) break;
        }
        
        if (canPlace) {
          // Check adjacency constraints
          const adjacencyScore = evaluateAdjacency(block, currentX, currentY, grid, buildingFeatures);
          
          if (adjacencyScore >= -1) { // Allow placement if not severely violating constraints
            // Place block
            for (let y = currentY; y < currentY + block.height && y < gridHeight; y++) {
              for (let x = currentX; x < currentX + block.width && x < gridWidth; x++) {
                grid[y][x] = block.id;
              }
            }
            
            block.x = currentX;
            block.y = currentY;
            block.placed = true;
            placedBlocks.push(block);
            
            maxRowHeight = Math.max(maxRowHeight, block.height);
            currentX += block.width + 1;
            placed = true;
          } else {
            currentX += 2;
          }
        } else {
          currentX += 1;
        }
        
        attempts++;
      }
    });
  });
  
  return { grid, placedBlocks, buildingFeatures };
};

const evaluateAdjacency = (block, x, y, grid, buildingFeatures) => {
  let score = 0;
  
  // Check priority 0 adjacencies (non-compromisable)
  const priority0Rules = ADJACENCY_RULES[0];
  priority0Rules.forEach(rule => {
    if (block.name === rule.from) {
      const nearbyFeature = buildingFeatures.find(f => 
        f.id.includes(rule.to.replace(/\s+/g, '-').toUpperCase())
      );
      
      if (nearbyFeature) {
        const distance = Math.sqrt(
          Math.pow(x - nearbyFeature.x, 2) + Math.pow(y - nearbyFeature.y, 2)
        );
        
        if (rule.relation === 'must_avoid' && distance < 10) {
          score += rule.weight * 5; // Strong penalty
        } else if (rule.relation === 'must_be_near' && distance < 15) {
          score += Math.abs(rule.weight); // Reward
        } else if (rule.relation === 'must_be_near' && distance > 30) {
          score -= 2; // Penalty for being too far
        }
      }
    }
  });
  
  // Check priority 2 adjacencies (block to building feature)
  const priority2Rules = ADJACENCY_RULES[2];
  priority2Rules.forEach(rule => {
    if (block.name === rule.from) {
      const nearbyFeature = buildingFeatures.find(f => 
        f.id.includes(rule.to.replace(/\s+/g, '-').toUpperCase()) || f.type === 'core'
      );
      
      if (nearbyFeature) {
        const distance = Math.sqrt(
          Math.pow(x - nearbyFeature.x, 2) + Math.pow(y - nearbyFeature.y, 2)
        );
        
        if (rule.relation === 'preferred_near' && distance < 20) {
          score += rule.weight;
        }
      }
    }
  });
  
  return score;
};

// ============ ANALYSIS & VALIDATION ============

const analyzeConstraints = (placedBlocks, superblocks, buildingFeatures) => {
  const analysis = {
    geometricRectangularity: 0,
    compositionMatches: [],
    adjacencyCompliance: { p0: 0, p1: 0, p2: 0, p3: 0, p4: 0, p5: 0 },
    totalBlocks: placedBlocks.length,
    placedBlocks: placedBlocks.length,
    violations: []
  };
  
  // 1. Geometric rectangularity
  let rectangularBlocks = 0;
  placedBlocks.forEach(block => {
    const aspectRatio = Math.max(block.width, block.height) / Math.min(block.width, block.height);
    if (aspectRatio <= 2) {
      rectangularBlocks++;
    }
  });
  analysis.geometricRectangularity = (rectangularBlocks / placedBlocks.length) * 100;
  
  // 2. Composition matches
  superblocks.forEach(sb => {
    analysis.compositionMatches.push({
      id: sb.id,
      match: sb.compositionMatch,
      target: sb.targetComposition,
      actual: sb.actualComposition
    });
  });
  
  // 3. Adjacency compliance
  let p0Compliant = 0, p0Total = 0;
  let p2Compliant = 0, p2Total = 0;
  let p3Compliant = 0, p3Total = 0;
  
  // Check Priority 0 (non-compromisable)
  ADJACENCY_RULES[0].forEach(rule => {
    const blocks = placedBlocks.filter(b => b.name === rule.from);
    blocks.forEach(block => {
      p0Total++;
      const nearbyFeature = buildingFeatures.find(f => 
        f.id.includes(rule.to.replace(/\s+/g, '-').toUpperCase())
      );
      
      if (nearbyFeature && block.x !== null) {
        const distance = Math.sqrt(
          Math.pow(block.x - nearbyFeature.x, 2) + Math.pow(block.y - nearbyFeature.y, 2)
        );
        
        if (rule.relation === 'must_avoid' && distance >= 10) {
          p0Compliant++;
        } else if (rule.relation === 'must_be_near' && distance < 15) {
          p0Compliant++;
        } else if (rule.relation === 'must_avoid' && distance < 10) {
          analysis.violations.push({
            priority: 0,
            block: block.id,
            rule: `${rule.from} must avoid ${rule.to}`,
            status: 'violated'
          });
        }
      }
    });
  });
  
  // Check Priority 2 (block to building feature)
  ADJACENCY_RULES[2].forEach(rule => {
    const blocks = placedBlocks.filter(b => b.name === rule.from);
    blocks.forEach(block => {
      p2Total++;
      const nearbyFeature = buildingFeatures.find(f => 
        f.id.includes(rule.to.replace(/\s+/g, '-').toUpperCase()) || f.type === 'core'
      );
      
      if (nearbyFeature && block.x !== null) {
        const distance = Math.sqrt(
          Math.pow(block.x - nearbyFeature.x, 2) + Math.pow(block.y - nearbyFeature.y, 2)
        );
        
        if (rule.relation === 'preferred_near' && distance < 20) {
          p2Compliant++;
        }
      }
    });
  });
  
  // Check Priority 3 (block to block)
  ADJACENCY_RULES[3].forEach(rule => {
    const fromBlocks = placedBlocks.filter(b => b.name === rule.from);
    const toBlocks = placedBlocks.filter(b => b.name === rule.to);
    
    fromBlocks.forEach(fromBlock => {
      toBlocks.forEach(toBlock => {
        if (fromBlock.x !== null && toBlock.x !== null) {
          p3Total++;
          const distance = Math.sqrt(
            Math.pow(fromBlock.x - toBlock.x, 2) + Math.pow(fromBlock.y - toBlock.y, 2)
          );
          
          if (rule.relation === 'should_avoid' && distance > 5) {
            p3Compliant++;
          }
        }
      });
    });
  });
  
  analysis.adjacencyCompliance.p0 = p0Total > 0 ? (p0Compliant / p0Total) * 100 : 100;
  analysis.adjacencyCompliance.p1 = analysis.geometricRectangularity; // Geometric resolution
  analysis.adjacencyCompliance.p2 = p2Total > 0 ? (p2Compliant / p2Total) * 100 : 100;
  analysis.adjacencyCompliance.p3 = p3Total > 0 ? (p3Compliant / p3Total) * 100 : 100;
  analysis.adjacencyCompliance.p4 = 100; // Superblock to superblock (not implemented in detail)
  analysis.adjacencyCompliance.p5 = 100; // Block to department (not implemented in detail)
  
  return analysis;
};

// ============ MAIN COMPONENT ============

export default function ZoningPlanner() {
  const [blocks, setBlocks] = useState([]);
  const [superblocks, setSuperblocks] = useState([]);
  const [gridData, setGridData] = useState(null);
  const [analysis, setAnalysis] = useState(null);
  const [hoveredCell, setHoveredCell] = useState(null);
  const [selectedBlock, setSelectedBlock] = useState(null);
  const [showAnalysis, setShowAnalysis] = useState(false);
  const canvasRef = useRef(null);
  
  useEffect(() => {
    initializeData();
  }, []);
  
  const initializeData = () => {
    const generatedBlocks = generateSyntheticBlocks(100);
    const generatedSuperblocks = generateSuperblocks(generatedBlocks);
    const { grid, placedBlocks, buildingFeatures } = placeBlocksOnGrid(generatedBlocks, generatedSuperblocks);
    const analysisResults = analyzeConstraints(placedBlocks, generatedSuperblocks, buildingFeatures);
    
    setBlocks(generatedBlocks);
    setSuperblocks(generatedSuperblocks);
    setGridData({ grid, placedBlocks, buildingFeatures });
    setAnalysis(analysisResults);
  };
  
  useEffect(() => {
    if (gridData && canvasRef.current) {
      drawGrid();
    }
  }, [gridData, hoveredCell]);
  
  const drawGrid = () => {
    const canvas = canvasRef.current;
    const ctx = canvas.getContext('2d');
    const cellSize = 8;
    
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    
    // Draw grid
    const { grid, buildingFeatures } = gridData;
    
    for (let y = 0; y < grid.length; y++) {
      for (let x = 0; x < grid[y].length; x++) {
        const cellId = grid[y][x];
        
        if (cellId) {
          // Determine color based on cell type
          let color = '#e0e0e0';
          
          const feature = buildingFeatures.find(f => f.id === cellId);
          if (feature) {
            if (feature.type === 'core') {
              color = '#1e3a8a'; // Dark blue for cores
            } else if (feature.type === 'anchor') {
              color = '#7c2d12'; // Brown for anchors
            }
          } else {
            const block = blocks.find(b => b.id === cellId);
            if (block) {
              switch (block.category) {
                case 'ME': color = '#3b82f6'; break; // Blue
                case 'WE': color = '#10b981'; break; // Green
                case 'US': color = '#f59e0b'; break; // Orange
                case 'SUPPORT': color = '#8b5cf6'; break; // Purple
                case 'SPECIALTY': color = '#ef4444'; break; // Red
              }
              
              // Highlight hovered block
              if (hoveredCell && hoveredCell.x === x && hoveredCell.y === y) {
                color = '#fbbf24'; // Yellow highlight
              }
              
              // Highlight selected block
              if (selectedBlock && selectedBlock.id === cellId) {
                color = '#fbbf24'; // Yellow highlight
              }
            }
          }
          
          ctx.fillStyle = color;
          ctx.fillRect(x * cellSize, y * cellSize, cellSize, cellSize);
          ctx.strokeStyle = '#ffffff';
          ctx.lineWidth = 0.5;
          ctx.strokeRect(x * cellSize, y * cellSize, cellSize, cellSize);
        }
      }
    }
  };
  
  const handleCanvasHover = (e) => {
    const canvas = canvasRef.current;
    const rect = canvas.getBoundingClientRect();
    const cellSize = 8;
    const x = Math.floor((e.clientX - rect.left) / cellSize);
    const y = Math.floor((e.clientY - rect.top) / cellSize);
    
    if (gridData && y < gridData.grid.length && x < gridData.grid[0].length) {
      const cellId = gridData.grid[y][x];
      if (cellId) {
        setHoveredCell({ x, y, id: cellId });
      } else {
        setHoveredCell(null);
      }
    }
  };
  
  const handleCanvasClick = (e) => {
    if (hoveredCell) {
      const block = blocks.find(b => b.id === hoveredCell.id);
      const feature = gridData.buildingFeatures.find(f => f.id === hoveredCell.id);
      
      if (block) {
        setSelectedBlock(block);
      } else if (feature) {
        setSelectedBlock(feature);
      }
    }
  };
  
  const getCategoryColor = (category) => {
    switch (category) {
      case 'ME': return 'bg-blue-500';
      case 'WE': return 'bg-green-500';
      case 'US': return 'bg-orange-500';
      case 'SUPPORT': return 'bg-purple-500';
      case 'SPECIALTY': return 'bg-red-500';
      default: return 'bg-gray-500';
    }
  };
  
  const getComplianceIcon = (percentage) => {
    if (percentage >= 80) return <CheckCircle className="w-5 h-5 text-green-500" />;
    if (percentage >= 60) return <AlertCircle className="w-5 h-5 text-yellow-500" />;
    return <XCircle className="w-5 h-5 text-red-500" />;
  };
  
  return (
    <div className="min-h-screen bg-gray-50 p-6">
      <div className="max-w-7xl mx-auto">
        <div className="bg-white rounded-lg shadow-lg p-6 mb-6">
          <h1 className="text-3xl font-bold text-gray-900 mb-2">Zoning Plan Generator</h1>
          <p className="text-gray-600 mb-4">
            Interactive 2D grid with adjacency-based block placement and constraint analysis
          </p>
          
          <div className="flex gap-4 mb-4">
            <button
              onClick={initializeData}
              className="px-4 py-2 bg-blue-600 text-white rounded hover:bg-blue-700 transition"
            >
              Regenerate Layout
            </button>
            <button
              onClick={() => setShowAnalysis(!showAnalysis)}
              className="px-4 py-2 bg-green-600 text-white rounded hover:bg-green-700 transition"
            >
              {showAnalysis ? 'Hide' : 'Show'} Analysis
            </button>
          </div>
          
          {/* Legend */}
          <div className="flex flex-wrap gap-4 mb-4 p-4 bg-gray-50 rounded">
            <div className="flex items-center gap-2">
              <div className="w-4 h-4 bg-blue-500"></div>
              <span className="text-sm">ME (Mission Essential)</span>
            </div>
            <div className="flex items-center gap-2">
              <div className="w-4 h-4 bg-green-500"></div>
              <span className="text-sm">WE (We-space)</span>
            </div>
            <div className="flex items-center gap-2">
              <div className="w-4 h-4 bg-orange-500"></div>
              <span className="text-sm">US (User Services)</span>
            </div>
            <div className="flex items-center gap-2">
              <div className="w-4 h-4 bg-purple-500"></div>
              <span className="text-sm">Support</span>
            </div>
            <div className="flex items-center gap-2">
              <div className="w-4 h-4 bg-red-500"></div>
              <span className="text-sm">Specialty</span>
            </div>
            <div className="flex items-center gap-2">
              <div className="w-4 h-4 bg-blue-900"></div>
              <span className="text-sm">Cores</span>
            </div>
            <div className="flex items-center gap-2">
              <div className="w-4 h-4 bg-amber-900"></div>
              <span className="text-sm">Anchors</span>
            </div>
          </div>
        </div>
        
        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          {/* Grid Visualization */}
          <div className="lg:col-span-2 bg-white rounded-lg shadow-lg p-6">
            <h2 className="text-xl font-bold mb-4">2D Floor Plan Grid</h2>
            <div className="overflow-auto border-2 border-gray-300 rounded">
              <canvas
                ref={canvasRef}
                width={800}
                height={640}
                onMouseMove={handleCanvasHover}
                onClick={handleCanvasClick}
                className="cursor-pointer"
              />
            </div>
            
            {hoveredCell && (
              <div className="mt-4 p-4 bg-blue-50 rounded border border-blue-200">
                <h3 className="font-semibold text-blue-900 mb-2">Hovered Cell Info</h3>
                <p className="text-sm text-gray-700">
                  <strong>ID:</strong> {hoveredCell.id}
                </p>
                {blocks.find(b => b.id === hoveredCell.id) && (
                  <>
                    <p className="text-sm text-gray-700">
                      <strong>Block:</strong> {blocks.find(b => b.id === hoveredCell.id).name}
                    </p>
                    <p className="text-sm text-gray-700">
                      <strong>Superblock:</strong> {blocks.find(b => b.id === hoveredCell.id).superblockId}
                    </p>
                  </>
                )}
              </div>
            )}
          </div>
          
          {/* Block Info Panel */}
          <div className="bg-white rounded-lg shadow-lg p-6">
            <h2 className="text-xl font-bold mb-4">Block Details</h2>
            
            {selectedBlock ? (
              <div className="space-y-3">
                <div className="p-3 bg-gray-50 rounded">
                  <p className="text-sm font-semibold text-gray-600">ID</p>
                  <p className="text-lg">{selectedBlock.id}</p>
                </div>
                
                {selectedBlock.name && (
                  <>
                    <div className="p-3 bg-gray-50 rounded">
                      <p className="text-sm font-semibold text-gray-600">Name</p>
                      <p className="text-sm">{selectedBlock.name}</p>
                    </div>
                    
                    <div className="p-3 bg-gray-50 rounded">
                      <p className="text-sm font-semibold text-gray-600">Category</p>
                      <div className="flex items-center gap-2 mt-1">
                        <div className={`w-3 h-3 ${getCategoryColor(selectedBlock.category)}`}></div>
                        <p>{selectedBlock.category}</p>
                      </div>
                    </div>
                    
                    <div className="p-3 bg-gray-50 rounded">
                      <p className="text-sm font-semibold text-gray-600">Dimensions</p>
                      <p>{selectedBlock.width}m × {selectedBlock.height}m ({selectedBlock.area}m²)</p>
                    </div>
                    
                    <div className="p-3 bg-gray-50 rounded">
                      <p className="text-sm font-semibold text-gray-600">Superblock</p>
                      <p>{selectedBlock.superblockId || 'N/A'}</p>
                    </div>
                    
                    <div className="p-3 bg-gray-50 rounded">
                      <p className="text-sm font-semibold text-gray-600">Position</p>
                      <p>X: {selectedBlock.x}, Y: {selectedBlock.y}</p>
                    </div>
                  </>
                )}
                
                {selectedBlock.type && (
                  <div className="p-3 bg-gray-50 rounded">
                    <p className="text-sm font-semibold text-gray-600">Type</p>
                    <p className="capitalize">{selectedBlock.type}</p>
                  </div>
                )}
              </div>
            ) : (
              <div className="text-center text-gray-500 py-8">
                <Info className="w-12 h-12 mx-auto mb-2 opacity-50" />
                <p>Click on a block to see details</p>
              </div>
            )}
          </div>
        </div>
        
        {/* Analysis Panel */}
        {showAnalysis && analysis && (
          <div className="mt-6 bg-white rounded-lg shadow-lg p-6">
            <h2 className="text-2xl font-bold mb-6">Constraint Analysis</h2>
            
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
              {/* Geometric Rectangularity */}
              <div className="p-4 bg-blue-50 rounded-lg border border-blue-200">
                <div className="flex items-center justify-between mb-2">
                  <h3 className="font-semibold text-blue-900">Geometric Rectangularity</h3>
                  {getComplianceIcon(analysis.geometricRectangularity)}
                </div>
                <p className="text-3xl font-bold text-blue-700">
                  {analysis.geometricRectangularity.toFixed(1)}%
                </p>
                <p className="text-sm text-gray-600 mt-1">
                  Blocks with aspect ratio ≤ 2:1
                </p>
              </div>
              
              {/* Priority 0 Compliance */}
              <div className="p-4 bg-red-50 rounded-lg border border-red-200">
                <div className="flex items-center justify-between mb-2">
                  <h3 className="font-semibold text-red-900">P0: Non-Compromisable</h3>
                  {getComplianceIcon(analysis.adjacencyCompliance.p0)}
                </div>
                <p className="text-3xl font-bold text-red-700">
                  {analysis.adjacencyCompliance.p0.toFixed(1)}%
                </p>
                <p className="text-sm text-gray-600 mt-1">
                  Critical adjacency rules followed
                </p>
              </div>
              
              {/* Priority 2 Compliance */}
              <div className="p-4 bg-green-50 rounded-lg border border-green-200">
                <div className="flex items-center justify-between mb-2">
                  <h3 className="font-semibold text-green-900">P2: Block-to-Feature</h3>
                  {getComplianceIcon(analysis.adjacencyCompliance.p2)}
                </div>
                <p className="text-3xl font-bold text-green-700">
                  {analysis.adjacencyCompliance.p2.toFixed(1)}%
                </p>
                <p className="text-sm text-gray-600 mt-1">
                  Building feature proximity
                </p>
              </div>
              
              {/* Priority 3 Compliance */}
              <div className="p-4 bg-yellow-50 rounded-lg border border-yellow-200">
                <div className="flex items-center justify-between mb-2">
                  <h3 className="font-semibold text-yellow-900">P3: Block-to-Block</h3>
                  {getComplianceIcon(analysis.adjacencyCompliance.p3)}
                </div>
                <p className="text-3xl font-bold text-yellow-700">
                  {analysis.adjacencyCompliance.p3.toFixed(1)}%
                </p>
                <p className="text-sm text-gray-600 mt-1">
                  Inter-block adjacencies
                </p>
              </div>
              
              {/* Placement Stats */}
              <div className="p-4 bg-purple-50 rounded-lg border border-purple-200">
                <h3 className="font-semibold text-purple-900 mb-2">Placement Stats</h3>
                <p className="text-3xl font-bold text-purple-700">
                  {analysis.placedBlocks}/{analysis.totalBlocks}
                </p>
                <p className="text-sm text-gray-600 mt-1">
                  Blocks successfully placed
                </p>
              </div>
              
              {/* Superblock Count */}
              <div className="p-4 bg-indigo-50 rounded-lg border border-indigo-200">
                <h3 className="font-semibold text-indigo-900 mb-2">Superblocks</h3>
                <p className="text-3xl font-bold text-indigo-700">
                  {superblocks.length}
                </p>
                <p className="text-sm text-gray-600 mt-1">
                  Total superblock clusters
                </p>
              </div>
            </div>
            
            {/* Superblock Composition Analysis */}
            <div className="mt-6">
              <h3 className="text-xl font-bold mb-4">Superblock Composition Analysis</h3>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                {analysis.compositionMatches.slice(0, 6).map(sb => (
                  <div key={sb.id} className="p-4 border border-gray-200 rounded-lg">
                    <div className="flex items-center justify-between mb-3">
                      <h4 className="font-semibold">{sb.id}</h4>
                      <div className="flex items-center gap-2">
                        {getComplianceIcon(sb.match)}
                        <span className="font-bold text-lg">{sb.match.toFixed(1)}%</span>
                      </div>
                    </div>
                    
                    <div className="space-y-2 text-sm">
                      <div className="flex justify-between">
                        <span className="text-gray-600">ME:</span>
                        <span>
                          <span className="text-blue-600 font-semibold">
                            {(sb.actual.ME * 100).toFixed(0)}%
                          </span>
                          <span className="text-gray-400 ml-1">
                            (target: {(sb.target.me * 100).toFixed(0)}%)
                          </span>
                        </span>
                      </div>
                      <div className="flex justify-between">
                        <span className="text-gray-600">WE:</span>
                        <span>
                          <span className="text-green-600 font-semibold">
                            {(sb.actual.WE * 100).toFixed(0)}%
                          </span>
                          <span className="text-gray-400 ml-1">
                            (target: {(sb.target.we * 100).toFixed(0)}%)
                          </span>
                        </span>
                      </div>
                      <div className="flex justify-between">
                        <span className="text-gray-600">US:</span>
                        <span>
                          <span className="text-orange-600 font-semibold">
                            {(sb.actual.US * 100).toFixed(0)}%
                          </span>
                          <span className="text-gray-400 ml-1">
                            (target: {(sb.target.us * 100).toFixed(0)}%)
                          </span>
                        </span>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            </div>
            
            {/* Violations */}
            {analysis.violations.length > 0 && (
              <div className="mt-6">
                <h3 className="text-xl font-bold mb-4 text-red-700">Constraint Violations</h3>
                <div className="space-y-2">
                  {analysis.violations.map((violation, idx) => (
                    <div key={idx} className="p-3 bg-red-50 border border-red-200 rounded flex items-start gap-3">
                      <XCircle className="w-5 h-5 text-red-500 flex-shrink-0 mt-0.5" />
                      <div>
                        <p className="font-semibold text-red-900">Priority {violation.priority}</p>
                        <p className="text-sm text-gray-700">{violation.block}: {violation.rule}</p>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
