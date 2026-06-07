#!/usr/bin/env python3
"""
Ekahau 报告审核工具 —— 自动检查 SSID 名称、信道配置、AP 部署、信号覆盖、图表颜色
用法：
    python3 ekahau_audit.py <Ekahau 报告.docx> [报告2.docx ...]
    python3 ekahau_audit.py /path/to/your_report.docx
    python3 ekahau_audit.py /path/to/your_report.docx --ssid MarriottBonvoy  # 指定期望的 SSID

功能：
  ✅ SSID 一致性检查         - 所有 AP 的广播 SSID 是否一致
  ✅ SSID 安全规范检查       - 敏感词、长度、格式
  ✅ 隐藏 SSID 检查          - 空 SSID / 隐藏网络统计
  ✅ 技术标准一致性检查      - 802.11ax/ac/n 是否统一
  ✅ 信道使用检查            - 2.4GHz 非重叠信道、5GHz DFS、同频干扰
  ✅ AP 配置完整性检查       - 双频、MAC 地址、厂商
  ✅ 信号覆盖强度检查        - 2.4G/5G 信号是否 ≥ -65 dBm
  ✅ SNR 信噪比检查          - SNR 是否 ≥ 25 dB
  ✅ 图表颜色校验            - 热力图颜色是否符合阈值设定
"""

import sys
import os
import re
from zipfile import ZipFile
from xml.etree import ElementTree
from collections import defaultdict

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}


# ============================================================
# 1. 解析 docx - 基于表格段落级别的精确解析
# ============================================================

def _get_paragraph_text(p_element):
    """从单个 w:p 元素提取文本"""
    texts = []
    for t in p_element.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t'):
        if t.text:
            texts.append(t.text)
    return ''.join(texts).strip()


def _get_cell_paragraphs(cell_element):
    """获取单元格内所有段落的文本列表"""
    paras = []
    for p in cell_element.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'):
        text = _get_paragraph_text(p)
        if text:
            paras.append(text)
    return paras


def _is_ap_name_line(text):
    """判断是否是 AP 名称行"""
    return bool(re.match(r'Name:\s*Measured\s+AP', text))


def _parse_ap_name(text):
    """从名称行解析 AP 信息"""
    m = re.search(r'Name:\s*(Measured\s+AP[_-][^(]+)\(?#?(\d+)\)?', text)
    if m:
        name = m.group(1).strip() + f' (#{m.group(2)})'
        ap_id = m.group(2)
        # 检查是否包含 Notes
        notes = ''
        notes_m = re.search(r'Notes[：:]\s*(.*)', text)
        if notes_m:
            notes = notes_m.group(1).strip()
        return {'name': name, 'id': ap_id, 'notes': notes}
    return None


def _parse_radio_paragraphs(paras):
    """从一组段落中解析射频和网络信息
    返回: [{'channel': int, 'radio_type': str, 'networks': [...]}]
    """
    radios = []
    current_channel = None
    current_radio_type = None
    current_networks = []

    for para in paras:
        # 检查是否包含 "Radio on channel: X"
        ch_match = re.search(r'Radio on channel:\s*(\d+)', para)
        if ch_match:
            # 保存上一个 radio（如果有）
            if current_channel is not None:
                radios.append({
                    'channel': current_channel,
                    'radio_type': current_radio_type,
                    'networks': current_networks,
                })

            current_channel = int(ch_match.group(1))
            current_radio_type = '2.4GHz' if 1 <= current_channel <= 13 else '5GHz'
            current_networks = []

            # 从同一行提取可能的网络信息
            for m in re.finditer(r'Mac:\s*([0-9a-f:]+),\s*SSID:\s*([^,]*),?\s*Technology:\s*(.+)', para):
                ssid = m.group(2).strip()
                current_networks.append({
                    'mac': m.group(1).lower(),
                    'ssid': ssid,
                    'technology': m.group(3).strip(),
                })
            continue

        # 检查是否是网络信息行
        net_match = re.search(r'Mac:\s*([0-9a-f:]+),\s*SSID:\s*([^,]*),?\s*Technology:\s*(.+)', para)
        if net_match and current_channel is not None:
            ssid = net_match.group(2).strip()
            current_networks.append({
                'mac': net_match.group(1).lower(),
                'ssid': ssid,
                'technology': net_match.group(3).strip(),
            })

    # 保存最后一个 radio
    if current_channel is not None:
        radios.append({
            'channel': current_channel,
            'radio_type': current_radio_type,
            'networks': current_networks,
        })

    return radios


def parse_ekahau_docx(filepath):
    """解析 Ekahau docx 报告"""

    with ZipFile(filepath) as z:
        xml_content = z.read('word/document.xml')
        root = ElementTree.fromstring(xml_content)

        # 提取文档所有段落（非表格部分）
        all_paragraphs = []
        for p in root.iter('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p'):
            text = _get_paragraph_text(p)
            if text:
                all_paragraphs.append(text)

        # 提取所有表格
        tables = root.findall('.//w:tbl', NS)

    # --- 解析元数据 ---
    metadata = {'project_name': '', 'location': '', 'responsible': '', 'floor': ''}
    for line in all_paragraphs:
        m = re.match(r'Name[：:]\s*(.+)', line)
        if m and not metadata['project_name']:
            metadata['project_name'] = m.group(1).strip()
        m = re.match(r'Location[：:]\s*(.+)', line)
        if m:
            metadata['location'] = m.group(1).strip()
        m = re.match(r'Responsible Person[：:]\s*(.+)', line)
        if m:
            metadata['responsible'] = m.group(1).strip()

    # 楼层检测
    for line in all_paragraphs:
        t = line.strip()
        if re.match(r'^B\d+$|^F\d+$|^L\d+$|^G$|^GF$', t):
            metadata['floor'] = t
            break

    # --- 解析 AP（基于表格结构） ---
    aps = []
    seen_ap_names = set()

    for tbl in tables:
        rows = tbl.findall('.//w:tr', NS)
        i = 0
        while i < len(rows):
            cells = rows[i].findall('.//w:tc', NS)
            cell0_paras = _get_cell_paragraphs(cells[0]) if cells else []

            # 查找 AP 名称行
            ap_info = None
            for para in cell0_paras:
                if _is_ap_name_line(para):
                    ap_info = _parse_ap_name(para)
                    break

            if ap_info and ap_info['name'] not in seen_ap_names:
                seen_ap_names.add(ap_info['name'])
                ap_entry = {
                    'name': ap_info['name'],
                    'id': ap_info['id'],
                    'notes': ap_info['notes'],
                    'radios': [],
                }

                # 如果 name 行本身没有 Notes，检查 cells[1]
                if not ap_entry['notes'] and len(cells) > 1:
                    notes_paras = [p for p in _get_cell_paragraphs(cells[1]) if p.startswith('Notes')]
                    for np_text in notes_paras:
                        nm = re.search(r'Notes[：:]\s*(.*)', np_text)
                        if nm:
                            ap_entry['notes'] = nm.group(1).strip()

                # 下一行是 radio 行
                if i + 1 < len(rows):
                    next_cells = rows[i + 1].findall('.//w:tc', NS)
                    if next_cells:
                        radio_paras = _get_cell_paragraphs(next_cells[0])
                        radios = _parse_radio_paragraphs(radio_paras)
                        ap_entry['radios'] = radios

                aps.append(ap_entry)
                i += 2  # 跳过 name 行和 radio 行
            else:
                i += 1

    # --- 如果表格未解析到 AP，尝试从段落文本解析 ---
    if not aps:
        aps = _parse_aps_from_paragraphs(all_paragraphs)

    return metadata, aps, all_paragraphs


def _parse_aps_from_paragraphs(paragraphs):
    """从段落文本格式解析 AP 数据（适用于无表格的报告版本）

    格式:
        Radio on channel: XX
        Mac: ..., SSID: ..., Technology: ...
        ... (4 条 MAC/radio)
        Radio on channel: YY
        ...
    """
    aps = []
    radios_buffer = []

    for line in paragraphs:
        ch_match = re.match(r'Radio on channel:\s*(\d+)', line)
        if ch_match:
            ch = int(ch_match.group(1))
            radio_type = '2.4GHz' if 1 <= ch <= 13 else '5GHz'
            radios_buffer.append({'channel': ch, 'radio_type': radio_type, 'networks': []})
        elif radios_buffer:
            net_match = re.search(r'Mac:\s*([0-9a-f:]+),\s*SSID:\s*([^,]*),?\s*Technology:\s*(.+)', line)
            if net_match:
                radios_buffer[-1]['networks'].append({
                    'mac': net_match.group(1).lower(),
                    'ssid': net_match.group(2).strip(),
                    'technology': net_match.group(3).strip(),
                })

    # 每 2 个 radio 为一组（2.4G + 5G）构成一个 AP
    for i in range(0, len(radios_buffer), 2):
        radios = radios_buffer[i:i + 2]
        if radios:
            ap_name = f'AP #{i // 2 + 1}'
            aps.append({
                'name': ap_name,
                'id': str(i // 2 + 1),
                'notes': '',
                'radios': radios,
            })

    return aps


# ============================================================
# 2. 审核规则
# ============================================================

class AuditRule:
    def __init__(self, name, description):
        self.name = name
        self.description = description

    def check(self, metadata, aps, paragraphs):
        raise NotImplementedError


class SSIDConsistencyCheck(AuditRule):
    """检查所有 AP 上的业务 SSID 是否一致"""

    def __init__(self, expected_ssid=None):
        super().__init__(
            'SSID 一致性检查',
            '验证所有 AP 的广播 SSID 是否一致、拼写是否正确'
        )
        self.expected_ssid = expected_ssid

    def check(self, metadata, aps, paragraphs):
        if not aps:
            return False, ['❌ 未发现任何 AP 信息']

        all_ssids = defaultdict(list)
        ap_ssid_coverage = defaultdict(set)  # ap_name -> set of ssids

        for ap in aps:
            for radio in ap.get('radios', []):
                for net in radio.get('networks', []):
                    ssid = net.get('ssid', '').strip()
                    label = f"{ap['name']} ({radio['radio_type']} Ch{radio['channel']})"
                    if ssid:
                        all_ssids[ssid].append(label)
                        ap_ssid_coverage[ap['name']].add(ssid)
                    else:
                        all_ssids['(隐藏/空 SSID)'].append(label)

        findings = []
        passed = True

        # 统计 SSID
        broadcast_ssids = {k: v for k, v in all_ssids.items() if k != '(隐藏/空 SSID)'}
        hidden_count = len(all_ssids.get('(隐藏/空 SSID)', []))

        if not broadcast_ssids:
            findings.append('⚠️  未发现任何广播 SSID（所有 SSID 均为空/隐藏）')
            return False, findings

        findings.append(f'📡 广播 SSID 列表:')
        for ssid, ap_list in broadcast_ssids.items():
            findings.append(f'   "{ssid}" → {len(ap_list)} 处')

        if hidden_count > 0:
            findings.append(f'🔇 隐藏 SSID（空 SSID）: {hidden_count} 处')

        # 指定 SSID 检查
        if self.expected_ssid:
            if self.expected_ssid not in broadcast_ssids:
                passed = False
                findings.append(f'❌ 期望的 SSID "{self.expected_ssid}" 未在任何 AP 上发现')
            else:
                covered_aps = set()
                for ap_name, ssids in ap_ssid_coverage.items():
                    if self.expected_ssid in ssids:
                        covered_aps.add(ap_name)
                total = len(aps)
                findings.append(
                    f'✅ "{self.expected_ssid}" 覆盖 {len(covered_aps)}/{total} 个 AP'
                    f' ({len(covered_aps)/total*100:.0f}%)'
                )
                if len(covered_aps) < total:
                    missing = [ap['name'] for ap in aps if ap['name'] not in covered_aps]
                    findings.append(f'⚠️  以下 AP 未广播 "{self.expected_ssid}":')
                    for m in missing:
                        findings.append(f'   - {m}')

        # 检查大小写
        if self.expected_ssid:
            for ssid in broadcast_ssids:
                if ssid.lower() == self.expected_ssid.lower() and ssid != self.expected_ssid:
                    passed = False
                    findings.append(f'❌ SSID 大小写不匹配: 期望 "{self.expected_ssid}", 实际 "{ssid}"')

        # 检查多个不同 SSID
        if len(broadcast_ssids) > 1:
            passed = False
            findings.append(f'❌ 发现 {len(broadcast_ssids)} 个不同的广播 SSID (可能存在测试/遗留网络)')

        return passed, findings


class SSIDSecurityCheck(AuditRule):
    def __init__(self):
        super().__init__('SSID 安全规范检查', '检查 SSID 是否包含敏感信息、符合命名规范')

    def check(self, metadata, aps, paragraphs):
        findings = []
        passed = True
        all_ssids = set()

        for ap in aps:
            for radio in ap.get('radios', []):
                for net in radio.get('networks', []):
                    ssid = net.get('ssid', '').strip()
                    if ssid:
                        all_ssids.add(ssid)

        for ssid in all_ssids:
            issues = []

            # 敏感词
            sensitive = ['admin', 'root', 'test', 'temp', 'cisco', 'huawei', 'aruba', 'default']
            for s in sensitive:
                if s in ssid.lower():
                    issues.append(f'包含敏感词 "{s}"')

            # 首尾空格
            if ssid != ssid.strip():
                issues.append('包含首尾空格')

            # 长度
            if len(ssid) > 32:
                issues.append(f'长度 {len(ssid)} 超过 32 字符限制')

            # 不可见字符
            if any(ord(c) < 32 for c in ssid):
                issues.append('包含控制字符')

            if issues:
                passed = False
                findings.append(f'⚠️  SSID "{ssid}": {"; ".join(issues)}')

        if passed and all_ssids:
            findings.append('✅ 所有 SSID 通过安全规范')

        return passed, findings


class HiddenSSIDCheck(AuditRule):
    def __init__(self):
        super().__init__('隐藏 SSID 检查', '统计空 SSID（隐藏网络）的使用情况')

    def check(self, metadata, aps, paragraphs):
        findings = []
        total = 0
        hidden = 0
        details = []

        for ap in aps:
            for radio in ap.get('radios', []):
                for net in radio.get('networks', []):
                    total += 1
                    if not net.get('ssid', '').strip():
                        hidden += 1
                        details.append(f"  {ap['name']} - {radio['radio_type']} Ch{radio['channel']} - {net['mac']}")

        ratio = hidden / total * 100 if total else 0

        if hidden == 0:
            findings.append('✅ 没有隐藏 SSID，所有网络均有名称')
            return True, findings

        findings.append(f'🔇 隐藏 SSID（空 SSID）: {hidden}/{total} ({ratio:.0f}%)')
        for d in details[:8]:
            findings.append(d)
        if len(details) > 8:
            findings.append(f'  ... 还有 {len(details) - 8} 个')

        if ratio > 50:
            findings.append('⚠️  隐藏 SSID 占比超过 50%，建议确认是否合理')
            return False, findings

        findings.append('ℹ️  隐藏 SSID 通常用于管理/访客/物联网网络')
        return True, findings


class ChannelUsageCheck(AuditRule):
    def __init__(self):
        super().__init__('信道使用检查', '检查 2.4GHz 重叠信道、5GHz DFS、同频干扰')

    def check(self, metadata, aps, paragraphs):
        findings = []
        passed = True
        ch_24 = []  # [(ap_name, channel)]
        ch_5 = []

        for ap in aps:
            for radio in ap.get('radios', []):
                ch = radio['channel']
                t = radio['radio_type']
                if t == '2.4GHz':
                    ch_24.append((ap['name'], ch))
                elif t == '5GHz':
                    ch_5.append((ap['name'], ch))

        # --- 2.4GHz ---
        if ch_24:
            used_ch = set(c for _, c in ch_24)
            non_overlap = {1, 6, 11}
            bad = [c for c in used_ch if c not in non_overlap]
            if bad:
                passed = False
                findings.append(f'❌ 2.4GHz 使用非标准信道: {sorted(bad)}')
                findings.append('💡 建议：2.4GHz 仅使用信道 1/6/11（不重叠）')
            else:
                findings.append(f'✅ 2.4GHz 使用标准非重叠信道: {sorted(used_ch)}')

            # 同频 AP 密度
            counts = defaultdict(list)
            for name, ch in ch_24:
                counts[ch].append(name)
            for ch, names in counts.items():
                if len(names) > 3:
                    findings.append(f'⚠️  2.4GHz 信道 {ch} 上有 {len(names)} 个 AP，可能同频干扰')

        # --- 5GHz ---
        if ch_5:
            used_ch_5 = sorted(set(c for _, c in ch_5))
            dfs = {52, 56, 60, 64, 100, 104, 108, 112, 116, 120, 124, 128, 132, 136, 140, 144}
            used_dfs = [c for c in used_ch_5 if c in dfs]
            if used_dfs:
                findings.append(f'ℹ️  5GHz 使用了 DFS 信道: {used_dfs}（可能受雷达影响跳频）')

            # 同频
            counts_5 = defaultdict(list)
            for name, ch in ch_5:
                counts_5[ch].append(name)
            for ch, names in counts_5.items():
                if len(names) > 2:
                    findings.append(f'⚠️  5GHz 信道 {ch} 上有 {len(names)} 个 AP，建议错开')
        else:
            findings.append('ℹ️  未发现 5GHz 配置')

        return passed, findings


class TechnologyConsistencyCheck(AuditRule):
    def __init__(self):
        super().__init__('技术标准一致性检查', '检查 Wi-Fi 标准（802.11ax/ac/n）是否统一')

    def check(self, metadata, aps, paragraphs):
        findings = []
        tech_map = defaultdict(set)
        tech_names = {
            '802.11ax': 'Wi-Fi 6',
            '802.11ac': 'Wi-Fi 5',
            '802.11n': 'Wi-Fi 4',
            '802.11be': 'Wi-Fi 7',
        }

        for ap in aps:
            for radio in ap.get('radios', []):
                for net in radio.get('networks', []):
                    tech = net.get('technology', '').strip()
                    if tech:
                        tech_map[tech].add(ap['name'])

        if not tech_map:
            findings.append('⚠️  未发现技术标准信息')
            return False, findings

        if len(tech_map) == 1:
            tech = list(tech_map.keys())[0]
            label = tech_names.get(tech, tech)
            findings.append(f'✅ 所有 AP 统一使用 {label}')
            return True, findings

        for tech, ap_set in tech_map.items():
            label = tech_names.get(tech, tech)
            findings.append(f'⚠️  {label}: {len(ap_set)} 个 AP')

        findings.append('💡 建议全场合规使用 Wi-Fi 6 (802.11ax)')
        return False, findings


class APConfigurationCheck(AuditRule):
    def __init__(self):
        super().__init__('AP 配置完整性检查', '检查 AP 双频支持、厂商分布、MAC 地址等')

    def check(self, metadata, aps, paragraphs):
        findings = []
        passed = True

        if not aps:
            return False, ['❌ 未发现 AP']

        findings.append(f'📊 共 {len(aps)} 个 AP')

        # 双频检查
        dual_band = 0
        single_24g = 0
        single_5g = 0
        for ap in aps:
            types = set(r['radio_type'] for r in ap.get('radios', []))
            if '2.4GHz' in types and '5GHz' in types:
                dual_band += 1
            elif '2.4GHz' in types:
                single_24g += 1
            elif '5GHz' in types:
                single_5g += 1

        if dual_band == len(aps):
            findings.append('✅ 所有 AP 支持双频（2.4GHz + 5GHz）')
        else:
            passed = False
            findings.append(f'⚠️ 双频 AP: {dual_band}, 仅 2.4G: {single_24g}, 仅 5G: {single_5g}')
            if single_24g:
                findings.append('💡 建议:仅 2.4G 的 AP 无法发挥 5GHz 性能优势')

        # 厂商分析
        oui_map = {
            '68:96:2e': 'Aruba/HPE',
            'e4:82:10': 'Huawei',
            '00:50:56': 'VMware',
            '00:0c:29': 'VMware',
            'a4:5e:60': 'Cisco',
            '70:3a:cb': 'Cisco',
            '00:1a:a1': 'Cisco',
            '38:11:25': 'Huawei',
            '64:6a:df': 'Huawei',
        }
        vendor_aps = defaultdict(set)

        for ap in aps:
            for radio in ap.get('radios', []):
                for net in radio.get('networks', []):
                    mac = net.get('mac', '')
                    if mac:
                        oui = mac[:8].lower()
                        vendor = oui_map.get(oui, '未知')
                        vendor_aps[vendor].add(ap['name'])

        if vendor_aps:
            parts = [f'{v}: {len(s)}个AP' for v, s in vendor_aps.items()]
            findings.append(f'🏭 厂商: {", ".join(parts)}')

        # 信道分布摘要
        ch_summary = defaultdict(lambda: {'count': 0, 'aps': []})
        for ap in aps:
            for radio in ap.get('radios', []):
                key = f"{radio['radio_type']} Ch{radio['channel']}"
                ch_summary[key]['count'] += 1
                ch_summary[key]['aps'].append(ap['name'])

        ch_items = sorted(ch_summary.items(), key=lambda x: (-x[1]['count'], x[0]))
        ch_detail = [f'{k} x{v["count"]}' for k, v in ch_items]
        findings.append(f'📡 信道分布: {", ".join(ch_detail)}')

        # 检查所有 AP 是否都有 notes
        no_notes = [ap['name'] for ap in aps if not ap.get('notes')]
        if no_notes and len(no_notes) == len(aps):
            findings.append('ℹ️  所有 AP 均无备注信息')

        return passed, findings


class SignalCoverageCheck(AuditRule):
    """增强信号覆盖检查 — 解析 2.4G/5G 信号数据并判断是否 ≥ -65dBm"""

    def __init__(self, threshold=-65):
        super().__init__(
            '信号覆盖强度检查 (-65dBm)',
            f'检查报告中 2.4G 和 5G 信号覆盖是否满足 ≥ {threshold} dBm'
        )
        self.threshold = threshold

    def check(self, metadata, aps, paragraphs):
        findings = []
        passed = True

        # 解析信号数据：{band: {avg, coverage_pct, ...}}
        signal_data = self._parse_signal_data(paragraphs)

        if not signal_data:
            # 无结构化数据时退回到关键字提取
            findings.append('ℹ️  未找到结构化信号数据，尝试从文本段落提取')
            seen = set()
            for line in paragraphs:
                for kw in ['Signal Strength', 'SNR', 'Coverage', '信号强度']:
                    if kw in line and line not in seen:
                        seen.add(line)
                        findings.append(f'📶 {line}')
                        break
            if not findings:
                findings.append('⚠️  无法从报告中自动提取信号强度数据')
                findings.append('💡 请在 Ekahau 报告中查看"Predicted Signal Strength"章节')
            findings.append(f'💡 标准: 信号 ≥ {self.threshold} dBm | SNR ≥ 25 dB')
            return False, findings

        # 逐频段检查
        for band in ['2.4', '5']:
            data = signal_data.get(band)
            if not data:
                findings.append(f'ℹ️  未找到 {band}GHz 频段的信号覆盖数据')
                continue

            avg = data.get('avg')
            cov_pct = data.get('coverage_pct')
            min_val = data.get('min_val')
            max_val = data.get('max_val')

            detail_parts = []
            if avg is not None:
                detail_parts.append(f'均值 {avg:.0f} dBm')
            if max_val is not None:
                detail_parts.append(f'最大 {max_val:.0f} dBm')
            if min_val is not None:
                detail_parts.append(f'最小 {min_val:.0f} dBm')
            if cov_pct is not None:
                detail_parts.append(f'≥{self.threshold}dBm 覆盖 {cov_pct:.1f}%')

            detail = ', '.join(detail_parts) if detail_parts else ''

            # 判定: 均值 >= threshold 或 覆盖率 >= 80%
            band_ok = False
            if avg is not None and avg >= self.threshold:
                band_ok = True
            if cov_pct is not None and cov_pct >= 80:
                band_ok = True

            if band_ok:
                findings.append(f'✅ {band}GHz 信号满足 ≥{self.threshold} dBm 标准 ({detail})')
            else:
                passed = False
                if avg is not None and avg < self.threshold:
                    findings.append(f'❌ {band}GHz 信号均值 {avg:.0f} dBm < {self.threshold} dBm ({detail})')
                elif cov_pct is not None and cov_pct < 80:
                    findings.append(f'⚠️  {band}GHz 频段 ≥{self.threshold}dBm 覆盖率仅 {cov_pct:.1f}% ({detail})')
                else:
                    findings.append(f'⚠️  {band}GHz 信号数据不足，无法判定 ({detail})')

        return passed, findings

    def _parse_signal_data(self, paragraphs):
        """从段落文本中解析 2.4G/5G 信号强度数据"""
        data = {}
        text = '\n'.join(paragraphs)

        # ============================================================
        # Pattern 1: Ekahau 标准格式 — "Predicted Signal Strength - X.X GHz"
        # 后跟覆盖百分比表格:
        #   >= -65 dBm     XX.X %
        #   Average: -XX dBm
        # ============================================================
        sections = re.split(
            r'(?:Predicted\s+)?Signal\s+Strength[\s\S]*?[-–]\s*(\d+[.]?\d*)\s*GHz',
            text, flags=re.IGNORECASE
        )
        # sections: [before, band1, content1, band2, content2, ...]
        if len(sections) >= 3:
            i = 1
            while i + 1 < len(sections):
                band = sections[i].strip()
                content = sections[i + 1]

                band_key = None
                if band.startswith('2'):
                    band_key = '2.4'
                elif band.startswith('5'):
                    band_key = '5'

                if band_key:
                    entry = self._parse_coverage_section(content)
                    if entry:
                        data[band_key] = entry

                i += 2

        # ============================================================
        # Pattern 2: 行内格式 — "2.4 GHz average signal: -XX dBm"
        # ============================================================
        for band_label, band_key in [('2.4', '2.4'), ('2.4 GHz', '2.4'), ('2.4GHz', '2.4'),
                                       ('5', '5'), ('5 GHz', '5'), ('5GHz', '5')]:
            if band_key in data:
                continue  # 已有数据则跳过
            # 匹配 "average signal: -XX dBm" 或 "avg: -XX dBm" 或 "Average: -XX dBm"
            avg_pattern = re.compile(
                rf'{re.escape(band_label)}.*?(?:average|avg|均值)[^:]*:?\s*(-?\d+)\s*dBm',
                re.IGNORECASE
            )
            avg_match = avg_pattern.search(text)
            if avg_match:
                entry = data.get(band_key, {})
                entry['avg'] = float(avg_match.group(1))
                data[band_key] = entry

        # ============================================================
        # Pattern 3: ">= -65 dBm: XX%" 或 ">= -65 dBm  XX%"
        # ============================================================
        # 先检查是否有覆盖率数据但没有 band key（Pattern 1 未匹配的情况）
        for band_label, band_key in [('2.4', '2.4'), ('2.4 GHz', '2.4'), ('2.4GHz', '2.4'),
                                       ('5', '5'), ('5 GHz', '5'), ('5GHz', '5')]:
            if band_key in data and 'coverage_pct' in data[band_key]:
                continue

            cov_pattern = re.compile(
                rf'{re.escape(band_label)}.*?(?:>=\s*|≥\s*)-?\d+\s*dBm[\s\S]*?(\d+[.]?\d*)\s*%',
                re.IGNORECASE
            )
            cov_match = cov_pattern.search(text)

            if not cov_match:
                # 尝试找 ">= -65 dBm" 紧跟百分比
                cov_pattern2 = re.compile(
                    rf'{re.escape(band_label)}.*?>=?\s*-?\d+\s*dBm\D*(\d+[.]?\d*)\s*%',
                    re.IGNORECASE
                )
                cov_match = cov_pattern2.search(text)

            if cov_match:
                entry = data.get(band_key, {})
                entry['coverage_pct'] = float(cov_match.group(1))
                data[band_key] = entry

        # ============================================================
        # Pattern 4: 独立覆盖率行 — ">= -65 dBm : XX %"
        # 需要判断前面最近出现的频段
        # ============================================================
        if '2.4' not in data or '5' not in data:
            lines = paragraphs
            current_band = None
            for line in lines:
                band_m = re.search(r'(2[.]?4|5)\s*GHz', line, re.IGNORECASE)
                if band_m:
                    raw = band_m.group(1)
                    current_band = '2.4' if raw.startswith('2') else '5'

                cov_m = re.search(r'(?:>=?|≥)\s*(-?\d+)\s*dBm[^:]*:?\s*(\d+[.]?\d*)\s*%', line)
                if cov_m and current_band and current_band not in data:
                    threshold_val = int(cov_m.group(1))
                    cov_pct = float(cov_m.group(2))
                    entry = data.get(current_band, {})
                    entry['coverage_pct'] = cov_pct
                    if threshold_val == abs(self.threshold):
                        entry['_exact_match'] = True
                    data[current_band] = entry

                # 也检查 "Average: -XX dBm"
                avg_m = re.search(r'(?:Average|Avg)[^:]*:?\s*(-?\d+)\s*dBm', line, re.IGNORECASE)
                if avg_m and current_band and current_band not in data:
                    entry = data.get(current_band, {})
                    entry['avg'] = float(avg_m.group(1))
                    data[current_band] = entry

        return data

    def _parse_coverage_section(self, content):
        """从覆盖率段落提取数据"""
        entry = {}

        # 找平均信号值
        avg_m = re.search(r'(?:Average|Avg|均值)[^:]*:?\s*(-?\d+)\s*dBm', content, re.IGNORECASE)
        if avg_m:
            entry['avg'] = float(avg_m.group(1))

        # 找 >= -65 dBm 的覆盖率
        cov_m = re.search(r'(?:>=?|≥)\s*(-?\d+)\s*dBm\D*(\d+[.]?\d*)\s*%', content)
        if cov_m:
            entry['coverage_pct'] = float(cov_m.group(2))

        # 找 min/max
        min_m = re.search(r'(?:Min|Minimum|最小)[^:]*:?\s*(-?\d+)\s*dBm', content, re.IGNORECASE)
        if min_m:
            entry['min_val'] = float(min_m.group(1))
        max_m = re.search(r'(?:Max|Maximum|最大)[^:]*:?\s*(-?\d+)\s*dBm', content, re.IGNORECASE)
        if max_m:
            entry['max_val'] = float(max_m.group(1))

        return entry if entry else None


class SNRTextCheck(AuditRule):
    """SNR 信噪比检查 — 解析报告中 SNR 覆盖数据，判断是否 ≥ 25dB"""

    def __init__(self, threshold=25):
        super().__init__(
            'SNR 信噪比检查 (≥25dB)',
            f'检查报告中 2.4G 和 5G 的 SNR 是否满足 ≥ {threshold} dB'
        )
        self.threshold = threshold

    def check(self, metadata, aps, paragraphs):
        findings = []
        passed = True

        snr_data = self._parse_snr_data(paragraphs)

        if not snr_data:
            findings.append('ℹ️  未在文本中找到 SNR 结构化数据')
            snr_lines = []
            for line in paragraphs:
                if re.search(r'SNR|Signal to Noise|信噪比', line, re.IGNORECASE):
                    if line not in snr_lines:
                        snr_lines.append(line)
            if snr_lines:
                findings.append('📶 找到 SNR 相关段落:')
                for line in snr_lines[:5]:
                    findings.append(f'  {line[:120]}')
            findings.append(f'💡 标准: SNR ≥ {self.threshold} dB')
            findings.append('💡 请查看报告中的"Signal to Noise Ratio"章节')
            return False, findings

        for band in ['2.4', '5']:
            data = snr_data.get(band)
            if not data:
                findings.append(f'ℹ️  未找到 {band}GHz 频段的 SNR 数据')
                continue

            avg = data.get('avg')
            cov_pct = data.get('coverage_pct')

            detail_parts = []
            if avg is not None:
                detail_parts.append(f'均值 {avg:.0f} dB')
            if cov_pct is not None:
                detail_parts.append(f'≥{self.threshold}dB 覆盖 {cov_pct:.1f}%')

            detail = ', '.join(detail_parts) if detail_parts else ''

            band_ok = False
            if avg is not None and avg >= self.threshold:
                band_ok = True
            if cov_pct is not None and cov_pct >= 80:
                band_ok = True

            if band_ok:
                findings.append(f'✅ {band}GHz SNR 满足 ≥{self.threshold} dB 标准 ({detail})')
            else:
                passed = False
                if avg is not None and avg < self.threshold:
                    findings.append(f'❌ {band}GHz SNR 均值 {avg:.0f} dB < {self.threshold} dB ({detail})')
                elif cov_pct is not None and cov_pct < 80:
                    findings.append(f'⚠️  {band}GHz ≥{self.threshold}dB 覆盖率仅 {cov_pct:.1f}% ({detail})')
                else:
                    findings.append(f'⚠️  {band}GHz SNR 数据不足，无法判定 ({detail})')

        return passed, findings

    def _parse_snr_data(self, paragraphs):
        text = '\n'.join(paragraphs)
        data = {}

        sections = re.split(
            r'(?:Predicted\s+)?(?:SNR|Signal\s+to\s+Noise\s+Ratio)[\s\S]*?[-–]\s*(\d+[.]?\d*)\s*GHz',
            text, flags=re.IGNORECASE
        )
        if len(sections) >= 3:
            i = 1
            while i + 1 < len(sections):
                band = sections[i].strip()
                content = sections[i + 1]
                band_key = '2.4' if band.startswith('2') else '5'
                entry = self._parse_snr_section(content)
                if entry:
                    data[band_key] = entry
                i += 2

        for band_label, band_key in [('2.4', '2.4'), ('2.4 GHz', '2.4'), ('2.4GHz', '2.4'),
                                       ('5', '5'), ('5 GHz', '5'), ('5GHz', '5')]:
            if band_key in data and 'avg' in data[band_key]:
                continue
            avg_pat = re.compile(
                rf'{re.escape(band_label)}.*?(?:average|avg|均值).*?(?:SNR|snr)[^:]*:?\s*(-?\d+)',
                re.IGNORECASE
            )
            avg_m = avg_pat.search(text)
            if avg_m:
                entry = data.get(band_key, {})
                entry['avg'] = float(avg_m.group(1))
                data[band_key] = entry

        if '2.4' not in data or '5' not in data:
            cur_band = None
            for line in paragraphs:
                bm = re.search(r'(2[.]?4|5)\s*GHz.*(?:SNR|snr)', line)
                if bm:
                    cur_band = '2.4' if bm.group(1).startswith('2') else '5'
                cm = re.search(r'(?:>=?|≥)\s*(\d+)\s*dB.*?(\d+[.]?\d*)\s*%', line)
                if cm and cur_band and cur_band not in data:
                    pct = float(cm.group(2))
                    entry = data.get(cur_band, {})
                    entry['coverage_pct'] = pct
                    data[cur_band] = entry
                am = re.search(r'(?:Average|Avg)[^:]*:?\s*(\d+[.]?\d*)\s*dB', line)
                if am and cur_band and cur_band not in data:
                    entry = data.get(cur_band, {})
                    entry['avg'] = float(am.group(1))
                    data[cur_band] = entry

        return data

    def _parse_snr_section(self, content):
        entry = {}
        avg_m = re.search(r'(?:Average|Avg|均值)[^:]*:?\s*(-?\d+)\s*dB', content, re.IGNORECASE)
        if avg_m:
            entry['avg'] = float(avg_m.group(1))
        cov_m = re.search(r'(?:>=?|≥)\s*(-?\d+)\s*dB\D*(\d+[.]?\d*)\s*%', content)
        if cov_m:
            entry['coverage_pct'] = float(cov_m.group(2))
        return entry if entry else None


class ChartColorCheck(AuditRule):
    """图表颜色校验 — 确认报告包含 VISUALIZATION STATISTICS 截图"""

    def __init__(self, filepath=None):
        super().__init__(
            '图表颜色校验',
            '确认 VISUALIZATION STATISTICS 截图已包含各章节覆盖数据'
        )
        self.filepath = filepath

    def set_filepath(self, filepath):
        self.filepath = filepath

    def check(self, metadata, aps, paragraphs):
        if not self.filepath:
            return False, ['❌ 未提供文件路径，无法分析图表']
        try:
            vs_images = self._extract_vs_images(self.filepath)
        except Exception as e:
            return False, [f'❌ 图表分析失败: {e}']
        if not vs_images:
            return False, ['ℹ️  未在报告中找到 VISUALIZATION STATISTICS 截图']
        findings, all_ok = [], True

        sec_map = {
            'SIGNAL_TOTAL': '📶 Signal Strength (total coverage)',
            'SIGNAL_SSID':  '📶 Signal Strength for SSID "MarriottBonvoy"',
            'SNR_SSID':     '📶 SNR for SSID "MarriottBonvoy"',
        }

        for sec_key, sec_label in sec_map.items():
            items = vs_images.get(sec_key, {})
            findings.append(f'{sec_label}')
            for band in ('2.4G', '5G'):
                img_data = items.get(band)
                if not img_data:
                    findings.append(f'  ⚠️ [{band}] 未检测到 VISUALIZATION STATISTICS 截图')
                    all_ok = False
                    continue

                # 分析 VISUALIZATION STATISTICS 截图中热力图区域的灰色不达标像素
                gray_pct = self._analyze_vs_gray(img_data['img'], img_data['size'])

                if gray_pct >= 3.0:
                    all_ok = False
                    findings.append(f'  ❌ [{band}] 存在 {gray_pct:.0f}% 灰色不达标区域（低于阈值）')
                else:
                    findings.append(f'  ✅ [{band}] 满足覆盖要求')
            findings.append('')

        return all_ok, findings

    def _analyze_vs_gray(self, img, size):
        """在 VISUALIZATION STATISTICS 截图的热力图区域内检测灰色不达标像素"""
        w, h = size
        rgba = img.convert('RGBA')
        pixels = list(rgba.getdata())
        rows = [pixels[i*w:(i+1)*w] for i in range(h)]

        # 步骤1：找热力图区域（绿色覆盖层密集的区域 = 楼层平面图）
        # 扫描找到绿色像素密集的列范围（排除右侧统计面板）
        green_cols = []
        for x in range(0, w, 4):
            green_count = 0
            for y in range(0, h, 4):
                px = rows[y][x]
                if len(px) >= 3:
                    r, g, b = px[0], px[1], px[2]
                    a = px[3] if len(px) > 3 else 255
                    if a < 128: continue
                    # 绿色热力图覆盖层
                    if g > r+15 and g > b+15 and g > 120:
                        green_count += 1
            if green_count > 5:  # 该列有足够多绿色像素
                green_cols.append(x)

        if not green_cols:
            # 没有绿色热力图，整图分析
            left, right, top, bottom = 0, w, 0, h
        else:
            # 绿色区域的范围（左侧热力图），排除右侧统计面板
            left = max(0, min(green_cols) - 20)
            right = min(w, max(green_cols) + 20)
            # 再缩右边 50px 排除色标尺的灰色段
            right = max(left, right - 50)
            # 也找绿色行范围
            # 也找绿色行范围
            green_rows = []
            for y in range(0, h, 4):
                for x in range(left, right, 4):
                    px = rows[y][x]
                    if len(px) >= 3:
                        r, g, b = px[0], px[1], px[2]
                        a = px[3] if len(px) > 3 else 255
                        if a < 128: continue
                        if g > r+15 and g > b+15 and g > 120:
                            green_rows.append(y)
                            break
            if green_rows:
                top = max(0, min(green_rows) - 10)
                bottom = min(h, max(green_rows) + 10)
            else:
                top, bottom = 0, h

        # 步骤2：在热力图区域内检测灰色像素
        total = 0
        gray = 0
        step = max(2, min(w, h) // 120)
        for y in range(top, bottom, step):
            for x in range(left, right, step):
                px = rows[y][x]
                if len(px) >= 3:
                    a = px[3] if len(px) > 3 else 255
                    if a < 128: continue
                    r, g, b = px[0], px[1], px[2]
                    if r > 240 and g > 240 and b > 240: continue
                    total += 1
                    cr = max(r,g,b) - min(r,g,b)
                    if cr < 25 and 80 < r < 190:
                        gray += 1

        if total == 0:
            return 0.0
        return 100.0 * gray / total

    def _extract_vs_images(self, filepath):
        """提取 VISUALIZATION STATISTICS 截图并分析热力图区域"""
        from zipfile import ZipFile
        from xml.etree import ElementTree
        from collections import defaultdict
        from PIL import Image
        import io, re

        NS_W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        NS_WP = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
        NS_R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

        with ZipFile(filepath) as z:
            # 读文档
            xml_content = z.read('word/document.xml')
            root = ElementTree.fromstring(xml_content)

            # 读关系映射
            rels_xml = z.read('word/_rels/document.xml.rels')
            rels_root = ElementTree.fromstring(rels_xml)
            rid_map = {}
            for rel in rels_root:
                rid = rel.get('Id')
                target = rel.get('Target')
                if rid and target:
                    if target.startswith('media/'):
                        rid_map[rid] = 'word/' + target
                    else:
                        rid_map[rid] = target

            # 遍历文档，跟踪章节和频段
            section = 'OTHER'
            band = ''
            section_media = []  # [(section, band, media_path)]

            def detect_section(text, current_section):
                lowered = text.lower()
                if lowered.startswith('name:') or lowered.startswith('mac:') or ', ssid:' in lowered:
                    return current_section
                if 'per ap' in lowered:
                    return 'PER_AP'
                has_band = bool(re.search(r'\b2[.]?\s*4', lowered) or re.search(r'\b5\s*g', lowered))
                quals = {'total', 'coverage', 'ssid', 'for', 'per', 'ap'}
                has_q = any(k in lowered for k in quals)
                if has_band and ('signal' in lowered or 'snr' in lowered) and not has_q and len(lowered) < 40:
                    return current_section
                if 'signal strength' in lowered and 'ssid' in lowered:
                    return 'SIGNAL_SSID'
                if 'snr' in lowered and 'ssid' in lowered:
                    return 'SNR_SSID'
                if 'signal strength' in lowered:
                    return 'SIGNAL_TOTAL'
                if 'snr' in lowered:
                    return 'SNR_SSID'
                return current_section

            def detect_band(text):
                lowered = text.lower()
                if re.search(r'\b2[.]?\s*4', lowered): return '2.4G'
                if re.search(r'\b5\s*g', lowered): return '5G'
                return band

            def process_para(para_elem):
                nonlocal section, band
                texts = []
                for t in para_elem.iter(NS_W + 't'):
                    if t.text: texts.append(t.text)
                text = ''.join(texts).strip()
                if text:
                    ns = detect_section(text, section)
                    if ns != section: section = ns
                    if section in ('SIGNAL_TOTAL','SIGNAL_SSID','SNR_SSID'):
                        nb = detect_band(text)
                        if nb: band = nb
                drawings = list(para_elem.iter(NS_WP + 'inline')) + list(para_elem.iter(NS_WP + 'anchor'))
                for dw in drawings:
                    for blip in dw.iter('{http://schemas.openxmlformats.org/drawingml/2006/main}blip'):
                        embed = blip.get(NS_R + 'embed') or blip.get(NS_R + 'link')
                        if embed and embed in rid_map:
                            section_media.append((section, band, rid_map[embed]))

            for elem in root.iter():
                tag = elem.tag
                if tag == NS_W + 'p': process_para(elem)
                elif tag == NS_W + 'tbl':
                    for p in elem.iter(NS_W + 'p'): process_para(p)

            # 去重，同组取非白色面积最大的（排除大量空白占位图）
            seen = {}
            for sec, b, path in section_media:
                key = (sec, b)
                # 快速评估图片内容：白度 < 60% 才算有效
                try:
                    data = z.read(path)
                    if len(data) < 10000: continue
                    temp_img = Image.open(io.BytesIO(data)).convert('RGB')
                    tw, th = temp_img.size
                    white_px = sum(1 for r,g,b in list(temp_img.getdata())[::100] if r>240 and g>240 and b>240)
                    white_pct = 100.0 * white_px / max(1, len(list(temp_img.getdata())[::100]))
                    if white_pct > 60:
                        continue  # 跳过大量空白图片
                    content_score = (100 - white_pct) * tw * th  # 非白色内容×总面积
                except:
                    content_score = z.getinfo(path).file_size if path in z.namelist() else 0
                if key not in seen or content_score > seen[key][1]:
                    seen[key] = (path, content_score)

            # 只保留目标章节，返回图片信息和图像数据
            result = defaultdict(dict)
            for (sec, b), (path, _) in seen.items():
                if sec not in ('SIGNAL_TOTAL', 'SIGNAL_SSID', 'SNR_SSID'):
                    continue
                data = z.read(path)
                if len(data) < 10000: continue
                try:
                    img = Image.open(io.BytesIO(data))
                    w, h = img.size
                    if w < 100 or h < 100: continue
                    result[sec][b] = {
                        'name': path.split('/')[-1],
                        'size': (w, h),
                        'section': sec,
                        'band': b,
                        'img': img.copy(),  # 保存图像供分析
                    }
                except:
                    continue

            return dict(result)

# ============================================================
# 3. 审核引擎
# ============================================================

def run_audit(filepath, expected_ssid=None):
    print('=' * 65)
    print('  Ekahau 无线报告审核工具')
    print('=' * 65)
    print(f'  文件: {filepath}')
    print()

    try:
        metadata, aps, paragraphs = parse_ekahau_docx(filepath)
    except Exception as e:
        print(f'❌ 解析失败: {e}')
        print('  请确认文件是 Ekahau 生成的 .docx 报告')
        return False

    print(f'📖 完成解析: {len(aps)} 个 AP')

    if metadata['project_name']:
        print()
        print('━' * 65)
        print('  项目信息')
        print('━' * 65)
        labels = {'project_name': '项目', 'location': '位置', 'responsible': '负责人', 'floor': '楼层'}
        for k, v in metadata.items():
            if v:
                print(f'  {labels.get(k, k)}: {v}')

    print()

    checks = [
        SSIDConsistencyCheck(expected_ssid),
        SSIDSecurityCheck(),
        HiddenSSIDCheck(),
        TechnologyConsistencyCheck(),
        ChannelUsageCheck(),
        APConfigurationCheck(),
        ChartColorCheck(filepath),  # 传文件路径用于图片分析
    ]

    all_passed = True
    summary = []

    for check in checks:
        print('━' * 65)
        print(f'  📋 {check.name}')
        print(f'     {check.description}')
        print('━' * 65)

        try:
            passed, findings = check.check(metadata, aps, paragraphs)
        except Exception as e:
            passed = False
            findings = [f'❌ 检查异常: {e}']

        for f in findings:
            print(f'  {f}')

        status = '✅ 通过' if passed else '⚠️  需关注'
        print(f'  → {status}')
        print()
        summary.append((check.name, passed))
        if not passed:
            all_passed = False

    print('=' * 65)
    print('  审核汇总')
    print('=' * 65)
    ok = sum(1 for _, p in summary if p)
    print(f'  通过 {ok}/{len(summary)}')

    if all_passed:
        print('  ✅ 全部通过！')
    else:
        print(f'  ⚠️  {len(summary) - ok} 项需关注')
    print('=' * 65)

    return all_passed


def run_audit_batch(filepaths, expected_ssid=None):
    """批量审核多个文件"""
    total_files = len(filepaths)
    passed_files = 0

    print('\n' + '#' * 65)
    print(f'  📂 批量审核: 共 {total_files} 个文件')
    print('#' * 65 + '\n')

    for i, fp in enumerate(filepaths, 1):
        print(f'\n{"─" * 65}')
        print(f'  [{i}/{total_files}] 正在审核: {os.path.basename(fp)}')
        print(f'{"─" * 65}\n')

        ok = run_audit(fp, expected_ssid)
        if ok:
            passed_files += 1

    print()
    print('=' * 65)
    print('  📊 批量审核汇总')
    print('=' * 65)
    print(f'  通过: {passed_files}/{total_files}')
    for fp in filepaths:
        print(f'    {os.path.basename(fp)}')
    if passed_files == total_files:
        print('  ✅ 全部文件通过审核！')
    else:
        print(f'  ⚠️  {total_files - passed_files} 个文件需关注')
    print('=' * 65)

    return passed_files == total_files


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    expected = None

    # 提取 --ssid 参数
    if '--ssid' in sys.argv:
        idx = sys.argv.index('--ssid')
        if idx + 1 < len(sys.argv):
            expected = sys.argv[idx + 1]
        # 从参数列表中移除 --ssid 及其值
        sys.argv = sys.argv[:idx] + sys.argv[idx+2:]

    # 剩余参数为文件路径
    filepaths = [a for a in sys.argv[1:] if not a.startswith('--')]

    if not filepaths:
        print(__doc__)
        sys.exit(1)

    # 检查文件扩展名
    for fp in filepaths:
        if not fp.endswith('.docx'):
            print(f'⚠️  警告: {os.path.basename(fp)} 不是 .docx 格式')

    if len(filepaths) == 1:
        run_audit(filepaths[0], expected)
    else:
        run_audit_batch(filepaths, expected)


import gradio as gr

with gr.Blocks() as demo:
    gr.Markdown("# Ekahau 报告审核")
    file = gr.File(label="上传 DOCX")
    output = gr.Textbox()
    btn = gr.Button("开始分析")
    btn.click(analyze_ekahau_docx, inputs=file, outputs=output)

if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 10000))
    demo.launch(server_name="0.0.0.0", server_port=port)
