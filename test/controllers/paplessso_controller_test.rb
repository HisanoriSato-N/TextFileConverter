require "test_helper"

class PaplesssoControllerTest < ActionDispatch::IntegrationTest
  test "should get top" do
    get paplessso_top_url
    assert_response :success
  end
end
