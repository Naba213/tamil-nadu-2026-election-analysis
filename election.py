#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      Lenovo
#
# Created:     29-05-2026
# Copyright:   (c) Lenovo 2026
# Licence:     <your licence>
#-------------------------------------------------------------------------------
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os

# ── folder path ──────────────────────────────────────────────
PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.join(PROJECT_DIR, "data")
OUT  = os.path.join(BASE, "outputs")
os.makedirs(OUT, exist_ok=True)

# ── load data ─────────────────────────────────────────────────
df21 = pd.read_csv(os.path.join(BASE, "tn_2021_results.csv"))
df26 = pd.read_csv(os.path.join(BASE, "tn_2026_results.csv"))
master = pd.read_csv(os.path.join(BASE, "constituency_master.csv"))

# ── remove NOTA ───────────────────────────────────────────────
df21 = df21[df21['party'] != 'NOTA']
df26 = df26[df26['party'] != 'NOTA']

# ── total votes per constituency ──────────────────────────────
total21 = df21.groupby('ac_number')['votes'].sum().reset_index()
total21.columns = ['ac_number', 'total_votes_21']

total26 = df26.groupby('ac_number')['votes'].sum().reset_index()
total26.columns = ['ac_number', 'total_votes_26']

# ── winner per constituency ───────────────────────────────────
win21 = df21.loc[df21.groupby('ac_number')['votes'].idxmax()][['ac_number','candidate','party','votes','region']]
win21.columns = ['ac_number','candidate_21','party_21','votes_21','region']

win26 = df26.loc[df26.groupby('ac_number')['votes'].idxmax()][['ac_number','candidate','party','votes','region']]
win26.columns = ['ac_number','candidate_26','party_26','votes_26','region']

# ── add total votes and vote share ────────────────────────────
win21 = win21.merge(total21, on='ac_number')
win21['share_21'] = (win21['votes_21'] / win21['total_votes_21'] * 100).round(2)

win26 = win26.merge(total26, on='ac_number')
win26['share_26'] = (win26['votes_26'] / win26['total_votes_26'] * 100).round(2)

# ── join 2021 and 2026 winners ────────────────────────────────
winners = win21.merge(win26, on='ac_number', suffixes=('_21','_26'))

print("✅ Winners table ready!")
print(winners[['ac_number','party_21','party_26','share_21','share_26','region_21']].head(10))
print("\nTotal constituencies:", len(winners))

# ════════════════════════════════════════════════════════════════
# Q1 — SEATS WON BY PARTY PER REGION (2026)
# ════════════════════════════════════════════════════════════════
seats_26 = win26.groupby(['region','party_26']).size().reset_index()
seats_26.columns = ['region','party','seats_2026']

# keep only top parties
top_parties = ['DMK','AIADMK','INC','TVK','PMK','VCK','BJP']
seats_26 = seats_26[seats_26['party'].isin(top_parties)]

pivot_q1 = seats_26.pivot(index='region', columns='party', values='seats_2026').fillna(0)

print("\n✅ Q1 — Seats per region per party (2026):")
print(pivot_q1)

fig, ax = plt.subplots(figsize=(12, 6))
pivot_q1.plot(kind='bar', ax=ax, colormap='tab10', width=0.8)
ax.set_title('Seats won by party across Tamil Nadu regions — 2026', fontsize=14, fontweight='bold')
ax.set_xlabel('Region')
ax.set_ylabel('Number of Seats')
ax.legend(title='Party', bbox_to_anchor=(1.05, 1))
ax.tick_params(axis='x', rotation=30)
plt.tight_layout()
plt.savefig(os.path.join(OUT, 'Q1_seats_by_region.png'), dpi=150)
plt.show()
print("✅ Q1 chart saved!")

# ════════════════════════════════════════════════════════════════
# Q2 — FLIP STORY
# ════════════════════════════════════════════════════════════════
winners['flipped'] = winners['party_21'] != winners['party_26']

total_flipped = winners['flipped'].sum()
print(f"\n✅ Q2 — Total flipped seats: {total_flipped} out of {len(winners)}")

flip_matrix = winners[winners['flipped']].groupby(['party_21','party_26']).size().reset_index()
flip_matrix.columns = ['from_party','to_party','count']
flip_matrix = flip_matrix.sort_values('count', ascending=False)
print("\nFlip matrix (top flows):")
print(flip_matrix.head(15))

pivot_flip = flip_matrix.pivot(index='from_party', columns='to_party', values='count').fillna(0)
fig, ax = plt.subplots(figsize=(10, 6))
sns.heatmap(pivot_flip, annot=True, fmt='.0f', cmap='YlOrRd', ax=ax, linewidths=0.5)
ax.set_title('Seat flips — From party (2021) → To party (2026)', fontsize=14, fontweight='bold')
ax.set_xlabel('Won by in 2026')
ax.set_ylabel('Held by in 2021')
plt.tight_layout()
plt.savefig(os.path.join(OUT, 'Q2_flip_heatmap.png'), dpi=150)
plt.show()
print("✅ Q2 chart saved!")

# ════════════════════════════════════════════════════════════════
# Q6 — MARGIN OF VICTORY
# ════════════════════════════════════════════════════════════════
avg_share_21 = win21['share_21'].mean().round(2)
avg_share_26 = win26['share_26'].mean().round(2)

above50_21 = (win21['share_21'] > 50).sum()
above50_26 = (win26['share_26'] > 50).sum()

below35_21 = (win21['share_21'] < 35).sum()
below35_26 = (win26['share_26'] < 35).sum()

print(f"\n✅ Q6 — Average winner vote share: 2021 = {avg_share_21}%  |  2026 = {avg_share_26}%")
print(f"Winners above 50%:  2021 = {above50_21}  |  2026 = {above50_26}")
print(f"Winners below 35%:  2021 = {below35_21}  |  2026 = {below35_26}")

fig, ax = plt.subplots(figsize=(10, 6))
ax.hist(win21['share_21'], bins=20, alpha=0.6, color='steelblue', label='2021', edgecolor='white')
ax.hist(win26['share_26'], bins=20, alpha=0.6, color='tomato', label='2026', edgecolor='white')
ax.axvline(35, color='gray', linestyle='--', linewidth=1, label='35% line')
ax.axvline(50, color='black', linestyle='--', linewidth=1, label='50% line')
ax.set_title('Winner vote share distribution — 2021 vs 2026', fontsize=14, fontweight='bold')
ax.set_xlabel('Vote Share %')
ax.set_ylabel('Number of Constituencies')
ax.legend()
plt.tight_layout()
plt.savefig(os.path.join(OUT, 'Q6_vote_share_histogram.png'), dpi=150)
plt.show()
print("✅ Q6 chart saved!")

print("\n🎉 ALL DONE! Check your outputs folder:")
print(OUT)

# Save winners table for Power BI
win26_save = win26.copy()
win26_save.columns = [str(c).replace('_26','').replace('_21','') if c not in ['ac_number','region'] else c for c in win26_save.columns]
win26_save.to_csv(os.path.join(OUT, 'winners_2026.csv'), index=False)

win21_save = win21.copy()
win21_save.columns = [str(c).replace('_21','').replace('_26','') if c not in ['ac_number','region'] else c for c in win21_save.columns]
win21_save.to_csv(os.path.join(OUT, 'winners_2021.csv'), index=False)

print("✅ Winners CSV saved!")
print("2026 columns:", win26_save.columns.tolist())
print("2021 columns:", win21_save.columns.tolist())