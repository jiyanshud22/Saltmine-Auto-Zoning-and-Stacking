# Documentation

## Overview

**Zoning Plan Generator** system that creates optimized floor plans by placing blocks on 2D grid while respecting multiple levels of adjacency constraints (Priority 0-5).

## Project Structure

```
zoning-plan-generator/
├── zoning_algorithm.py          # Core Python algorithm
├── visualize_zoning.py           # Python visualization (matplotlib)
├── zoning_planner.jsx            # React interactive visualization
├── zoning_data.json              # Generated synthetic data
├── blocks.csv                    # Block details export
├── grid.csv                      # Grid layout export
├── zoning_plan.png               # Main floor plan visualization
├── superblock_analysis.png       # Superblock composition charts
└── compliance_dashboard.png      # Constraint compliance metrics
```

## Features

### 1. Synthetic Data Generation
- **100+ blocks** with realistic dimensions based on category
- **12 superblocks** with target compositions (ME: 50-65%, WE: 20-30%, etc.)
- Automatic block assignment to superblocks maintaining proper ratios

### 2. Priority-Based Adjacency Constraints

#### Priority 0: Non-Compromisable Adjacencies
- **Critical constraints that MUST be followed**
- Examples:
  - IDF/Server Room must avoid AHU
  - IDF/Server Room must avoid windows/glazing
  - Server Room must avoid toilets
  - Restaurant must be near main entry
  - Bike storage must be near service entry

#### Priority 1: Geometrical Resolution
- Maximizes rectangular block shapes
- Target: Aspect ratio ≤ 2:1
- Affects: All blocks

#### Priority 2: Block-to-Building Feature
- **Preferred adjacencies to building elements**
- Examples:
  - IDF/Server Room preferred near cores
  - Restaurant preferred near windows
  - Coffee bar preferred near windows

#### Priority 3: Block-to-Block Adjacencies
- Inter-block spacing requirements
- Examples:
  - Phone booths should avoid work desks
  - Lockers should avoid work desks

#### Priority 4: Superblock-to-Superblock
- Framework for inter-superblock relationships

#### Priority 5: Block-to-Department
- Framework for department-level adjacencies

### 3. Block Categories

| Category | Description | Color | Examples |
|----------|-------------|-------|----------|
| **ME** | Mission Essential | Blue | Work desks, focus rooms, touchdown seats |
| **WE** | We-space (Collaboration) | Green | Phone booths, meeting rooms, huddle rooms |
| **US** | User Services | Orange | Restaurant, coffee bar, flex space |
| **SUPPORT** | Support Spaces | Purple | Lockers, storage, copy/print |
| **SPECIALTY** | Specialty Rooms | Red | IDF, server room, security |

### 4. Building Features

- **Cores**: CORE-A, CORE-B, CORE-C (Dark Blue)
- **Anchors**: Main Entry, Service Entry, Fire Egress (Brown)
- **Special Features**: AHU units, Toilets (Gray)

## Installation & Setup

### Prerequisites
```bash
# Python requirements
pip install numpy pandas matplotlib --break-system-packages

# For React visualization
# Requires Node.js environment or claude.ai artifacts
```

### Running the Python Algorithm

```bash
# Generate zoning plan and data
python zoning_algorithm.py

# Generate visualizations
python visualize_zoning.py
```

## Usage Guide

### Python Algorithm

The main algorithm (`zoning_algorithm.py`) performs the following steps:

1. **Generate Synthetic Blocks**
   - Creates 100 blocks with realistic dimensions
   - Categories: ME, WE, US, SUPPORT, SPECIALTY
   - Dimensions vary by category

2. **Create Superblocks**
   - Groups blocks into 12 superblocks
   - Each superblock has target composition ratios
   - Actual composition calculated and matched

3. **Place Blocks on Grid**
   - 100x80 meter grid
   - Places building features first (cores, anchors)
   - Places blocks by superblock
   - Evaluates adjacency constraints during placement

4. **Analyze Constraints**
   - Geometric rectangularity score
   - Adjacency compliance by priority (P0-P5)
   - Superblock composition matches
   - Violation detection

5. **Export Data**
   - JSON: Complete data export
   - CSV: Blocks and grid layouts
   - PNG: Visualizations

### React Visualization

The React component (`zoning_planner.jsx`) provides:

- **Interactive 2D Grid**: Canvas-based visualization
- **Hover Information**: Shows block details on hover
- **Click Selection**: Detailed block information panel
- **Analysis Dashboard**: Real-time constraint compliance
- **Regenerate**: Create new random layouts

### Key Functions

#### Python

```python
# Generate blocks
generator = SyntheticDataGenerator(config)
blocks = generator.generate_blocks(count=100)

# Create superblocks
superblocks = generator.generate_superblocks(blocks)

# Place on grid
placement_engine = PlacementEngine(grid_width=100, grid_height=80)
placed_blocks = placement_engine.place_blocks(blocks, superblocks)

# Analyze
analyzer = AnalysisEngine(blocks, superblocks, building_features, rules)
analysis = analyzer.generate_report()
```

#### React

```javascript
// Initialize data
const blocks = generateSyntheticBlocks(100);
const superblocks = generateSuperblocks(blocks);
const { grid, placedBlocks } = placeBlocksOnGrid(blocks, superblocks);

// Analyze
const analysis = analyzeConstraints(placedBlocks, superblocks);
```

## Algorithm Details

### Placement Strategy

1. **Grid Initialization**: Create 100x80 empty grid
2. **Feature Placement**: Place cores and anchors at fixed positions
3. **Superblock Iteration**: Process each superblock sequentially
4. **Block Sorting**: Sort blocks by priority (ME → WE → US → Support → Specialty)
5. **Position Search**: For each block:
   - Start at superblock's designated area
   - Check if space is available
   - Evaluate adjacency score
   - Place if score threshold met
   - Move to next position if not

### Adjacency Scoring

```python
score = 0

# For each applicable adjacency rule:
if rule.relation == 'must_avoid':
    if distance < threshold:
        score += rule.weight * penalty_multiplier
        
elif rule.relation == 'must_be_near':
    if distance < threshold:
        score += abs(rule.weight) * reward_multiplier
        
elif rule.relation == 'preferred_near':
    if distance < threshold:
        score += rule.weight
```

### Composition Matching

```python
# Calculate difference between target and actual
error = sum(|target[category] - actual[category]|)

# Convert to match percentage
match = max(0, 1 - error) * 100
```

## Output Files

### 1. zoning_data.json
Complete data export including:
- All blocks with positions
- Superblock compositions
- Building features
- Analysis results

### 2. blocks.csv
Tabular block data:
```csv
id,name,category,width,height,area,superblock_id,placed,x,y
BLK-001,Individual Zone Work Desks,ME,3,3,9,SB-01,True,15,20
```

### 3. grid.csv
Grid layout (100x80):
- Each cell contains block ID or null
- Can be imported to spreadsheet tools

### 4. zoning_plan.png
Main visualization showing:
- All placed blocks (color-coded)
- Building features (cores, anchors)
- Grid layout
- Legend

### 5. superblock_analysis.png
Pie charts showing:
- Composition breakdown for each superblock
- ME/WE/US/Support/Specialty percentages
- Match scores

### 6. compliance_dashboard.png
Metrics dashboard:
- Adjacency compliance by priority
- Placement statistics
- Geometric rectangularity
- Superblock matches

## Analysis Metrics

### Geometric Rectangularity
**Target**: 100%
- Percentage of blocks with aspect ratio ≤ 2:1
- Higher is better
- Affects visual quality and space efficiency

### Adjacency Compliance
**Target**: ≥80% for each priority
- P0 (Non-Compromisable): Most critical
- P1 (Geometric): Shape quality
- P2 (Block-to-Feature): Building element proximity
- P3 (Block-to-Block): Inter-block spacing
- P4/P5: Superblock and department level

### Superblock Composition Match
**Target**: ≥80%
- How well actual composition matches target
- Calculated per superblock
- Affects functional balance

### Placement Rate
**Target**: ≥80%
- Percentage of blocks successfully placed
- Limited by grid space and constraints
- Higher indicates better algorithm efficiency

## Customization

### Modify Block Types

Edit `ZoningConfig.BLOCK_TYPES`:
```python
BLOCK_TYPES = {
    'ME': ['Custom Block Type 1', 'Custom Block Type 2'],
    # ... add your block types
}
```

### Add Adjacency Rules

Edit `ZoningConfig.get_adjacency_rules()`:
```python
rules.append(AdjacencyRule(
    priority=0,
    from_block='Your Block',
    to_block='Target Feature',
    relation='must_avoid',  # or 'must_be_near', 'preferred_near', 'should_avoid'
    weight=-2  # -2 to 1
))
```

### Change Grid Size

```python
placement_engine = PlacementEngine(
    grid_width=120,   # Adjust width
    grid_height=100   # Adjust height
)
```

### Modify Superblock Compositions

Edit `ZoningConfig.SUPERBLOCK_COMPOSITIONS`:
```python
SUPERBLOCK_COMPOSITIONS = [
    {'me': 0.70, 'we': 0.20, 'us': 0.05, 'support': 0.05, 'specialty': 0.0},
    # Add more composition templates
]
```

## Troubleshooting

### Low Placement Rate (<70%)
- **Cause**: Too many blocks for available space
- **Solution**: Reduce block count or increase grid size

### Low P0 Compliance
- **Cause**: Difficult to satisfy all critical constraints
- **Solution**: Adjust building feature positions or relax threshold

### Poor Composition Matches
- **Cause**: Not enough blocks of certain categories
- **Solution**: Generate more balanced block distribution

### Blocks Overlapping
- **Cause**: Bug in placement algorithm
- **Solution**: Check `can_place_block()` function
