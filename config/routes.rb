Rails.application.routes.draw do
  get 'paplessso/top'
  get 'converts/download/:id',to: "converts#download",as: "download_file"
  resources :converts
  get '*not_found' => 'application#routing_error'
  post '*not_found' => 'application#routing_error'
end
