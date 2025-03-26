/**
 * Excel参数依赖关系可视化
 * 使用D3.js实现参数之间的依赖关系可视化
 */

// 全局变量
let allParameters = {};         // 所有参数
let dependencyData = {};        // 依赖关系数据
let selectedNodeId = null;      // 当前选中的节点ID
let svg = null;                 // SVG对象
let simulation = null;          // 力导向模拟
let nodes = [];                 // 节点数据
let links = [];                 // 连接线数据
let displayMode = 'all';        // 显示模式：all-所有参数, dependencies-仅依赖
let calculatedValues = {};      // 计算结果

// 初始化
$(document).ready(function() {
    // 加载参数数据
    loadParameters();
    
    // 绑定视图切换按钮
    $('#view-all').click(function() {
        $(this).addClass('active');
        $('#view-dependencies').removeClass('active');
        displayMode = 'all';
        updateVisualization();
    });
    
    $('#view-dependencies').click(function() {
        $(this).addClass('active');
        $('#view-all').removeClass('active');
        displayMode = 'dependencies';
        updateVisualization();
    });
    
    // 绑定计算按钮
    $('#calculation-form').on('submit', function(e) {
        e.preventDefault();
        calculateParameters();
    });
    
    // 实时计算：当输入参数变化时自动计算
    $(document).on('change', '.param-input', function() {
        calculateParameters();
    });
});

// 加载参数数据
function loadParameters() {
    console.log('开始加载参数数据...');
    $.ajax({
        url: '/api/parameters',
        type: 'GET',
        dataType: 'json',
        success: function(data) {
            console.log('参数数据加载成功:', data);
            allParameters = data;
            
            // 确保参数数据格式一致性
            normalizeParameterData(allParameters);
            
            // 渲染参数列表
            renderParameterLists(data);
            
            // 加载依赖关系
            loadDependencies();
        },
        error: function(xhr, status, error) {
            console.error('加载参数数据失败:', xhr.responseText);
            console.error('状态码:', xhr.status);
            console.error('错误信息:', error);
            
            let errorMsg = '加载参数数据失败';
            if (xhr.responseJSON && xhr.responseJSON.error) {
                errorMsg += ': ' + xhr.responseJSON.error;
            }
            
            alert(errorMsg);
            $('#visualization-container').html('<div class="alert alert-danger p-5 text-center">' + errorMsg + '<br>请返回重新上传文件</div>');
        }
    });
}

// 生成输入参数表单
function generateInputForm(params) {
    // 这个函数不再需要，因为输入表单已经在renderParameterLists中创建
    // 保留这个函数是为了向后兼容，但不执行任何操作
}

// 计算参数
function calculateParameters() {
    // 清除错误信息
    $('#calc-error').addClass('d-none');
    
    // 收集输入参数值
    const inputValues = {};
    $('.param-input').each(function() {
        const paramId = $(this).data('param-id');
        const value = $(this).val();
        
        if (value) {
            inputValues[paramId] = parseFloat(value);
        }
    });
    
    // 发送计算请求
    $.ajax({
        url: '/api/calculate',
        type: 'POST',
        contentType: 'application/json',
        data: JSON.stringify(inputValues),
        dataType: 'json',
        success: function(response) {
            // 保存计算结果
            calculatedValues = response.calculated_values;
            
            // 更新可视化中的值
            updateVisualizedValues();
            
            // 更新参数列表显示
            updateParameterLists();
            
            // 如果当前有选中的参数，更新其详情
            if (selectedNodeId && calculatedValues[selectedNodeId]) {
                updateParameterDetails(selectedNodeId);
            }
        },
        error: function(xhr) {
            console.error('计算参数值失败:', xhr.responseText);
            let errorMessage = '计算失败，请稍后重试。';
            if (xhr.responseJSON && xhr.responseJSON.error) {
                errorMessage = xhr.responseJSON.error;
            }
            $('#calc-error').removeClass('d-none').text(errorMessage);
        }
    });
}

// 更新可视化中的参数值
function updateVisualizedValues() {
    // 更新节点数据
    if (svg) {
        svg.selectAll('.node').each(function(d) {
            if (calculatedValues[d.id]) {
                const value = calculatedValues[d.id].value;
                d.value = value;
            }
        });
    }
}

// 更新参数列表
function updateParameterLists() {
    // 清除所有错误信息
    $('.param-error').addClass('d-none').text('');
    
    // 更新输入参数值
    allParameters.input_params.forEach(function(param, index) {
        if (calculatedValues[param.标识符]) {
            allParameters.input_params[index].值 = calculatedValues[param.标识符].value;
            // 更新输入参数表格中的值
            $(`#input-params-table tr[data-param-id="${param.标识符}"] input`).val(calculatedValues[param.标识符].value);
        }
    });
    
    // 更新中间参数值
    allParameters.intermediate_params.forEach(function(param, index) {
        if (calculatedValues[param.标识符]) {
            allParameters.intermediate_params[index].值 = calculatedValues[param.标识符].value;
            
            // 更新中间参数表格中的值
            const valueCell = $(`#intermediate-params-table tr[data-param-id="${param.标识符}"] input`);
            valueCell.val(formatParameterValue(calculatedValues[param.标识符].value, param.单位));
            
            // 检查是否有错误
            if (calculatedValues[param.标识符].error) {
                const errorDiv = $(`#intermediate-params-table tr[data-param-id="${param.标识符}"] .param-error`);
                errorDiv.text(calculatedValues[param.标识符].error).removeClass('d-none');
                // 对于有错误的单元格，添加错误样式
                valueCell.addClass('is-invalid');
            } else {
                // 移除错误样式
                valueCell.removeClass('is-invalid');
            }
        }
    });
    
    // 更新输出参数值
    allParameters.output_params.forEach(function(param, index) {
        if (calculatedValues[param.标识符]) {
            // 保存原始值，不做任何转换
            allParameters.output_params[index].值 = calculatedValues[param.标识符].value;
            
            // 更新输出参数表格中的值
            const valueCell = $(`#output-params-table tr[data-param-id="${param.标识符}"] input`);
            // 使用格式化函数处理值，但字符串直接显示
            valueCell.val(formatParameterValue(calculatedValues[param.标识符].value, param.单位));
            
            // 检查是否有错误
            if (calculatedValues[param.标识符].error) {
                const errorDiv = $(`#output-params-table tr[data-param-id="${param.标识符}"] .param-error`);
                errorDiv.text(calculatedValues[param.标识符].error).removeClass('d-none');
                // 对于有错误的单元格，添加错误样式
                valueCell.addClass('is-invalid');
            } else {
                // 移除错误样式
                valueCell.removeClass('is-invalid');
            }
        }
    });
}

// 更新参数详情
function updateParameterDetails(paramId) {
    if (paramId && calculatedValues[paramId]) {
        // 如果参数详情已打开，更新其值
        $('#param-details .param-value').text(calculatedValues[paramId].value !== null ? calculatedValues[paramId].value : '无');
    }
}

// 标准化参数数据格式，确保前端处理统一
function normalizeParameterData(data) {
    // 遍历所有参数类别
    ['input_params', 'intermediate_params', 'output_params', 'independent_params'].forEach(category => {
        if (data[category]) {
            data[category].forEach(param => {
                // 确保关键属性存在
                if (!param.hasOwnProperty('标识符')) {
                    param.标识符 = param.id || param.名称;
                }
                if (!param.hasOwnProperty('依赖')) {
                    param.依赖 = param.dependencies || [];
                }
                if (!param.hasOwnProperty('依赖描述')) {
                    param.依赖描述 = param.dependency_names || [];
                }
                if (!param.hasOwnProperty('公式描述')) {
                    param.公式描述 = param.formula_description || param.公式 || '';
                }
                
                // 确保值类型正确
                if (typeof param.值 === 'undefined' || param.值 === null) {
                    param.值 = 0;
                }
            });
        }
    });
}

// 加载依赖关系
function loadDependencies() {
    $.ajax({
        url: '/api/dependencies',
        type: 'GET',
        dataType: 'json',
        success: function(data) {
            dependencyData = data;
            
            // 初始化可视化
            initVisualization();
        },
        error: function(xhr) {
            console.error('加载依赖关系失败:', xhr.responseText);
            alert('加载依赖关系数据失败，请刷新页面重试。');
        }
    });
}

// 渲染参数列表
function renderParameterLists(data) {
    // 清空现有表格
    $('#input-params-table, #intermediate-params-table, #output-params-table').empty();
    
    // 渲染输入参数表格
    data.input_params.forEach(function(param) {
        $('#input-params-table').append(`
            <tr data-param-id="${param.标识符}" class="input-param">
                <td class="param-label">
                    <a href="#" onclick="showParameterDetails('${param.标识符}'); return false;">${param.名称}</a>
                    <small class="text-muted">${param.单位 ? '(' + param.单位 + ')' : ''}</small>
                </td>
                <td class="param-input-cell">
                    <input type="number" class="form-control form-control-sm param-input" 
                           data-param-id="${param.标识符}" value="${param.值 !== null ? param.值 : ''}" step="any">
                </td>
            </tr>
        `);
    });
    
    // 渲染中间参数表格
    data.intermediate_params.forEach(function(param) {
        $('#intermediate-params-table').append(`
            <tr data-param-id="${param.标识符}" class="intermediate-param">
                <td class="param-label">
                    <a href="#" onclick="showParameterDetails('${param.标识符}'); return false;">${param.名称}</a>
                    <small class="text-muted">${param.单位 ? '(' + param.单位 + ')' : ''}</small>
                </td>
                <td class="param-input-cell">
                    <input type="text" class="form-control form-control-sm" 
                           value="${formatParameterValue(param.值, param.单位)}" readonly>
                    <div class="param-error d-none"></div>
                </td>
            </tr>
        `);
    });
    
    // 渲染输出参数表格
    data.output_params.forEach(function(param) {
        $('#output-params-table').append(`
            <tr data-param-id="${param.标识符}" class="output-param">
                <td class="param-label">
                    <a href="#" onclick="showParameterDetails('${param.标识符}'); return false;">${param.名称}</a>
                    <small class="text-muted">${param.单位 ? '(' + param.单位 + ')' : ''}</small>
                </td>
                <td class="param-input-cell">
                    <input type="text" class="form-control form-control-sm" 
                           value="${formatParameterValue(param.值, param.单位)}" readonly>
                    <div class="param-error d-none"></div>
                </td>
            </tr>
        `);
    });
}

// 格式化参数值，添加单位和控制小数位
function formatParameterValue(value, unit) {
    if (value === null || value === undefined) {
        return '';
    }
    
    // 如果是字符串且包含特殊格式（如"1:4.52"），直接返回
    if (typeof value === 'string' && (value.includes(':') || value.includes('：'))) {
        return value;
    }
    
    // 尝试将值转换为数字
    let numValue = parseFloat(value);
    if (isNaN(numValue)) {
        return value; // 如果不是数字，直接返回原值
    }
    
    // 根据数值大小控制小数位数
    let formattedValue;
    if (Math.abs(numValue) >= 100) {
        // 对于较大的数，保留0-2位小数
        formattedValue = numValue.toFixed(Math.min(2, countDecimals(numValue)));
    } else if (Math.abs(numValue) >= 10) {
        // 中等大小的数，保留最多3位小数
        formattedValue = numValue.toFixed(Math.min(3, countDecimals(numValue)));
    } else if (Math.abs(numValue) >= 1) {
        // 小于10的数，保留最多4位小数
        formattedValue = numValue.toFixed(Math.min(4, countDecimals(numValue)));
    } else if (numValue === 0) {
        // 如果是0，不显示小数
        formattedValue = '0';
    } else {
        // 非常小的数，保留最多6位有效数字
        formattedValue = numValue.toPrecision(6);
    }
    
    // 去除末尾的0和不必要的小数点
    formattedValue = formattedValue.replace(/\.0+$/, '').replace(/(\.\d*[1-9])0+$/, '$1');
    
    return formattedValue;
}

// 计算数字的小数位数
function countDecimals(value) {
    if (Math.floor(value) === value) return 0;
    return value.toString().split('.')[1].length || 0;
}

// 初始化可视化
function initVisualization() {
    // 准备可视化数据
    prepareVisualizationData();
    
    // 创建SVG容器
    const container = d3.select('#visualization-container');
    const width = container.node().getBoundingClientRect().width;
    const height = container.node().getBoundingClientRect().height;
    
    svg = container.append('svg')
        .attr('width', width)
        .attr('height', height)
        .append('g')
        .attr('transform', 'translate(' + width / 2 + ',' + height / 2 + ')');
    
    // 添加缩放功能
    const zoom = d3.zoom()
        .scaleExtent([0.1, 4])
        .on('zoom', (event) => {
            svg.attr('transform', event.transform);
        });
    
    d3.select('#visualization-container svg')
        .call(zoom);
    
    // 创建力导向图
    simulation = d3.forceSimulation(nodes)
        .force('link', d3.forceLink(links).id(d => d.id).distance(100))
        .force('charge', d3.forceManyBody().strength(-300))
        .force('center', d3.forceCenter(0, 0))
        .force('collide', d3.forceCollide(30))
        .on('tick', ticked);
    
    // 创建连接线的容器
    svg.append('g')
        .attr('class', 'links');
        
    // 创建节点的容器
    svg.append('g')
        .attr('class', 'nodes');
    
    // 绘制连接线
    svg.select('.links')
        .selectAll('line')
        .data(links)
        .enter().append('line')
        .attr('class', 'link')
        .attr('stroke-width', 1.5);
    
    // 使用新的方法构建节点
    rebuildAllNodes(nodes);
    
    // 点击背景时取消高亮
    svg.on('click', function() {
        clearHighlights();
        selectedNodeId = null;
    });
    
    // 力导向图tick函数
    function ticked() {
        svg.selectAll('.link')
            .attr('x1', d => d.source.x)
            .attr('y1', d => d.source.y)
            .attr('x2', d => d.target.x)
            .attr('y2', d => d.target.y);
        
        svg.selectAll('.node')
            .attr('transform', d => `translate(${d.x}, ${d.y})`);
    }
}

// 准备可视化数据
function prepareVisualizationData() {
    nodes = [];
    links = [];
    
    // 添加节点
    // 输入参数
    allParameters.input_params.forEach(param => {
        nodes.push({
            id: param.标识符,
            name: param.名称,
            type: 'input',
            value: param.值,
            unit: param.单位
        });
    });
    
    // 中间参数
    allParameters.intermediate_params.forEach(param => {
        nodes.push({
            id: param.标识符,
            name: param.名称,
            type: 'intermediate',
            value: param.值,
            unit: param.单位
        });
    });
    
    // 输出参数
    allParameters.output_params.forEach(param => {
        nodes.push({
            id: param.标识符,
            name: param.名称,
            type: 'output',
            value: param.值,
            unit: param.单位
        });
    });
    
    // 独立参数
    allParameters.independent_params.forEach(param => {
        nodes.push({
            id: param.标识符,
            name: param.名称,
            type: 'independent',
            value: param.值,
            unit: param.单位
        });
    });
    
    // 添加连接线
    dependencyData.forEach(dep => {
        links.push({
            source: dep.target_id,  // 注意：这里是反向的，从依赖指向被依赖
            target: dep.source_id,
            value: 1
        });
    });
}

// 更新可视化
function updateVisualization() {
    // 根据显示模式过滤节点和连接线
    let filteredNodes = [];
    let filteredLinks = [];
    
    if (displayMode === 'all') {
        // 显示所有参数
        filteredNodes = nodes;
        filteredLinks = links;
    } else if (displayMode === 'dependencies' && selectedNodeId) {
        // 仅显示当前选中参数的依赖链
        const dependencyChain = new Set();
        collectDependencyChain(selectedNodeId, dependencyChain);
        dependencyChain.add(selectedNodeId);
        
        filteredNodes = nodes.filter(node => dependencyChain.has(node.id));
        filteredLinks = links.filter(link => 
            dependencyChain.has(link.source.id || link.source) && 
            dependencyChain.has(link.target.id || link.target)
        );
    } else {
        // 默认显示所有
        filteredNodes = nodes;
        filteredLinks = links;
    }
    
    // 更新力导向图
    simulation.nodes(filteredNodes);
    simulation.force('link').links(filteredLinks);
    simulation.alpha(1).restart();
    
    // 更新链接线
    const link = svg.select('.links').selectAll('.link')
        .data(filteredLinks, d => `${d.source.id || d.source}-${d.target.id || d.target}`);
    
    link.exit().remove();
    
    const linkEnter = link.enter().append('line')
        .attr('class', 'link');
        
    // 重建所有节点 - 这个部分解决了中文编码问题
    rebuildAllNodes(filteredNodes);
    
    // 如果有选中的节点，高亮其依赖链
    if (selectedNodeId) {
        highlightDependencyChain(selectedNodeId);
    }
}

// 完全重建所有节点的函数
function rebuildAllNodes(nodeData) {
    // 清除所有现有节点
    svg.select('.nodes').selectAll('*').remove();
    
    // 创建新节点
    const nodes = svg.select('.nodes')
        .selectAll('.node')
        .data(nodeData, d => d.id)
        .enter().append('g')
        .attr('class', d => `node ${getNodeClass(d)}`)
        .call(d3.drag()
            .on('start', function(event) {
                if (!event.active) simulation.alphaTarget(0.3).restart();
                event.subject.fx = event.subject.x;
                event.subject.fy = event.subject.y;
            })
            .on('drag', function(event) {
                event.subject.fx = event.x;
                event.subject.fy = event.y;
            })
            .on('end', function(event) {
                if (!event.active) simulation.alphaTarget(0);
                event.subject.fx = null;
                event.subject.fy = null;
            }));
    
    // 添加圆形
    nodes.append('circle')
        .attr('r', 10);
    
    // 添加文本标签，专门设置字体以解决中文问题
    nodes.append('text')
        .attr('dx', 12)
        .attr('dy', '.35em')
        .attr('font-family', "'Microsoft YaHei', 'SimHei', Arial, sans-serif")
        .text(d => d.name);
    
    // 添加点击事件
    nodes.on('click', function(event, d) {
        event.stopPropagation();
        showParameterDetails(d.id);
        highlightDependencyChain(d.id);
    });
}

// 获取节点CSS类
function getNodeClass(node) {
    switch(node.type) {
        case 'input':
            return 'input-param';
        case 'intermediate':
            return 'intermediate-param';
        case 'output':
            return 'output-param';
        default:
            return '';
    }
}

// 收集依赖链(递归)
function collectDependencyChain(nodeId, chain, visited = new Set()) {
    // 检测循环依赖
    if (visited.has(nodeId)) {
        return;  // 如果已经访问过，则停止递归
    }
    
    // 标记当前节点为已访问
    visited.add(nodeId);
    
    // 查找该节点的所有依赖
    links.forEach(link => {
        if ((link.target.id || link.target) === nodeId) {
            const sourceId = link.source.id || link.source;
            chain.add(sourceId);
            collectDependencyChain(sourceId, chain, visited);
        }
    });
}

// 高亮依赖链
function highlightDependencyChain(nodeId) {
    clearHighlights();
    selectedNodeId = nodeId;
    
    // 收集依赖链
    const dependencyChain = new Set();
    collectDependencyChain(nodeId, dependencyChain);
    
    // 高亮节点 - 延迟执行以确保节点已经完全渲染
    setTimeout(() => {
        svg.selectAll('.node')
            .classed('highlighted', d => dependencyChain.has(d.id) || d.id === nodeId);
        
        // 高亮链接线
        svg.selectAll('.link')
            .classed('highlighted', d => {
                const sourceId = d.source.id || d.source;
                const targetId = d.target.id || d.target;
                return (dependencyChain.has(sourceId) && dependencyChain.has(targetId)) ||
                       (dependencyChain.has(sourceId) && targetId === nodeId) ||
                       (sourceId === nodeId && dependencyChain.has(targetId));
            });
        
        // 高亮参数列表中的项
        $('.list-group-item').removeClass('active');
        $(`.list-group-item[data-param-id="${nodeId}"]`).addClass('active');
    }, 50);
}

// 清除所有高亮
function clearHighlights() {
    svg.selectAll('.node').classed('highlighted', false);
    svg.selectAll('.link').classed('highlighted', false);
    $('.list-group-item').removeClass('active');
}

// 显示参数详情
function showParameterDetails(paramId) {
    $.ajax({
        url: `/api/parameter_details/${paramId}`,
        type: 'GET',
        dataType: 'json',
        success: function(data) {
            // 更新数据中的值为计算后的最新值
            if (calculatedValues[paramId]) {
                data.value = calculatedValues[paramId].value;
            }
            
            renderParameterDetails(data);
            selectedNodeId = paramId;
            
            // 高亮依赖链
            highlightDependencyChain(paramId);
            
            // 如果是"仅显示依赖"模式，更新可视化
            if (displayMode === 'dependencies') {
                updateVisualization();
            }
        },
        error: function(xhr) {
            console.error('获取参数详情失败:', xhr.responseText);
            $('#param-details').html('<div class="alert alert-danger">获取参数详情失败</div>');
        }
    });
}

// 渲染参数详情
function renderParameterDetails(data) {
    $('#param-detail-title').text(data.name);
    
    // 确定参数类型
    let paramType = '';
    let typeClass = '';
    for (let i = 0; i < allParameters.input_params.length; i++) {
        if (allParameters.input_params[i].标识符 === data.id) {
            paramType = '输入参数';
            typeClass = 'input-param-badge';
            break;
        }
    }
    if (!paramType) {
        for (let i = 0; i < allParameters.intermediate_params.length; i++) {
            if (allParameters.intermediate_params[i].标识符 === data.id) {
                paramType = '中间参数';
                typeClass = 'intermediate-param-badge';
                break;
            }
        }
    }
    if (!paramType) {
        for (let i = 0; i < allParameters.output_params.length; i++) {
            if (allParameters.output_params[i].标识符 === data.id) {
                paramType = '输出参数';
                typeClass = 'output-param-badge';
                break;
            }
        }
    }
    
    // 格式化显示值
    const formattedValue = formatParameterValue(data.value, data.unit);
    
    // 检查是否有计算错误
    let errorMessage = '';
    if (calculatedValues[data.id] && calculatedValues[data.id].error) {
        errorMessage = calculatedValues[data.id].error;
    }
    
    let html = `
        <div class="row">
            <div class="col-md-4">
                <div class="card">
                    <div class="card-header">基本信息</div>
                    <div class="card-body">
                        <p><strong>名称:</strong> ${data.name} <span class="badge ${typeClass}">${paramType}</span></p>
                        <p><strong>单位:</strong> ${data.unit || '无'}</p>
                        <p><strong>当前值:</strong> <span class="param-value calculation-result">${formattedValue || '无'}</span></p>
                        ${errorMessage ? `<div class="alert alert-danger mt-2 p-2">${errorMessage}</div>` : ''}
                    </div>
                </div>
            </div>
            <div class="col-md-8">
                <div class="card">
                    <div class="card-header">计算公式</div>
                    <div class="card-body">
    `;
    
    if (data.formula) {
        let formula = data.formula_description || data.formula;
        let displayFormula = formula;
        
        // 将公式中的参数名转为链接
        if (data.dependency_names && data.dependencies) {
            // 准备用于显示的公式（用于显示，包含彩色链接）
            data.dependency_names.forEach((depName, index) => {
                if (index < data.dependencies.length) {
                    const depId = data.dependencies[index];
                    if (depId) {
                        // 检查是否为重命名参数（包含下划线）
                        const pattern = new RegExp(`\\b${escapeRegExp(depName)}\\b`, 'g');
                        displayFormula = displayFormula.replace(pattern, 
                            `<a href="#" class="param-reference" onclick="showParameterDetails('${depId}'); return false;">${depName}</a>`);
                    }
                }
            });
            
            // 美化原始公式（替换^为上标，替换*为×等）
            let beautifiedFormula = formula
                .replace(/\^(\d+)/g, '<sup>$1</sup>')  // 将^2转换为上标
                .replace(/\*/g, ' × ')                 // 将*转换为乘号
                .replace(/\//g, ' ÷ ')                 // 将/转换为除号
                .replace(/\+/g, ' + ')                 // 添加+号周围的空格
                .replace(/\-/g, ' - ');                // 添加-号周围的空格
            
            // 只显示计算公式（替换后的公式）
            html += `
                <div class="param-formula-section">
                    <div class="formula-title">计算公式：</div>
                    <div class="param-formula">${displayFormula}</div>
                </div>`;
        } else {
            html += `<div class="param-formula">${displayFormula}</div>`;
        }
    } else {
        html += '<p class="text-muted">该参数没有计算公式</p>';
    }
    
    html += `
                    </div>
                </div>
                
                <div class="card mt-3">
                    <div class="card-header">依赖参数</div>
                    <div class="card-body">
    `;
    
    // 检查是否存在循环依赖
    let hasCyclicDependency = false;
    
    // 从后端直接获取循环依赖标记
    if (data.has_circular_dependency) {
        hasCyclicDependency = true;
    } else {
        // 备用方法：检查依赖链中是否有循环
        if (data.dependency_chain) {
            const checkForCycles = (nodes) => {
                if (!nodes) return false;
                for (let node of nodes) {
                    if (node.is_cycle) {
                        hasCyclicDependency = true;
                        return true;
                    }
                    if (checkForCycles(node.children)) {
                        return true;
                    }
                }
                return false;
            };
            checkForCycles(data.dependency_chain);
        }
    }
    
    // 如果存在循环依赖，显示警告
    if (hasCyclicDependency) {
        html += `
            <div class="alert alert-warning circular-dependency-warning">
                <i class="fas fa-exclamation-triangle"></i> 
                检测到循环依赖关系。这可能会导致计算问题，请检查您的公式。
            </div>
        `;
    }
    
    if (data.dependency_names && data.dependency_names.length > 0) {
        html += '<div class="table-responsive"><table class="table table-sm"><thead><tr><th>参数名称</th><th>当前值</th></tr></thead><tbody>';
        data.dependency_names.forEach((depName, index) => {
            if (index < data.dependencies.length) {
                const depId = data.dependencies[index];
                
                // 获取依赖参数的详细信息
                let depValue = '';
                let depUnit = '';
                let depType = '';
                
                if (calculatedValues[depId]) {
                    depValue = formatParameterValue(calculatedValues[depId].value);
                    
                    // 查找单位和类型
                    for (const paramList of [allParameters.input_params, allParameters.intermediate_params, allParameters.output_params]) {
                        for (const param of paramList) {
                            if (param.标识符 === depId) {
                                depUnit = param.单位 || '';
                                if (paramList === allParameters.input_params) {
                                    depType = 'input-param-badge';
                                } else if (paramList === allParameters.intermediate_params) {
                                    depType = 'intermediate-param-badge';
                                } else {
                                    depType = 'output-param-badge';
                                }
                                break;
                            }
                        }
                        if (depType) break;
                    }
                }
                
                html += `
                    <tr>
                        <td>
                            <a href="#" onclick="showParameterDetails('${depId}'); return false;" class="${depType ? 'text-'+depType.replace('-badge', '') : ''}">${depName}</a>
                            ${depUnit ? `<small class="text-muted">(${depUnit})</small>` : ''}
                        </td>
                        <td>
                            <span class="badge badge-light">${depValue || '未计算'}</span>
                        </td>
                    </tr>
                `;
            }
        });
        html += '</tbody></table></div>';
    } else {
        html += '<p class="text-muted">该参数没有依赖其他参数</p>';
    }
    
    html += `
                    </div>
                </div>
            </div>
        </div>
    `;
    
    $('#param-details').html(html);
}

// 转义正则表达式中的特殊字符
function escapeRegExp(string) {
    return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); // $& 表示匹配到的子字符串
}

// 根据参数名获取参数ID
function getDependencyId(name, dependencies) {
    for (let i = 0; i < allParameters.input_params.length; i++) {
        if (allParameters.input_params[i].名称 === name) {
            return allParameters.input_params[i].标识符;
        }
    }
    
    for (let i = 0; i < allParameters.intermediate_params.length; i++) {
        if (allParameters.intermediate_params[i].名称 === name) {
            return allParameters.intermediate_params[i].标识符;
        }
    }
    
    for (let i = 0; i < allParameters.output_params.length; i++) {
        if (allParameters.output_params[i].名称 === name) {
            return allParameters.output_params[i].标识符;
        }
    }
    
    // 如果在各个参数组中都找不到，直接返回依赖数组中相应的ID
    const index = allParameters.dependency_names.indexOf(name);
    return index !== -1 ? dependencies[index] : null;
}

// 在前端接收和显示值时添加特殊处理
function displayValue(value) {
  if (typeof value === 'string' && value.includes(':')) {
    // 确保完整显示带冒号的值
    return value;
  }
  // 其他正常处理...
} 