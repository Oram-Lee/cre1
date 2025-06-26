// updateStats 함수 수정 - 탭 카운트도 함께 업데이트
function updateStats() {
    const validBuildings = buildingsData.filter(b => b.lat && b.lng);
    const displayedValidBuildings = currentDisplayedBuildings.filter(b => b.lat && b.lng);
    
    document.getElementById('stat-total').textContent = buildingsData.length;
    document.getElementById('stat-displayed').textContent = displayedValidBuildings.length;
    document.getElementById('stat-geocoded').textContent = validBuildings.length;
    
    // 탭 카운트 업데이트 추가
    if (typeof updateTabCounts === 'function') {
        updateTabCounts();
    }
}

// updateCartCount 함수 제거 (탭 카운트로 대체)
function updateCartCount() {
    // 기존 cart-count 엘리먼트가 있는 경우를 위한 호환성 유지
    const cartCount = document.getElementById('cart-count');
    if (cartCount) {
        cartCount.textContent = selectedBuildings.length;
    }
    
    // 탭 카운트 업데이트
    if (typeof updateTabCounts === 'function') {
        updateTabCounts();
    }
}
